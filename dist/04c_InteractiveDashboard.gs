// ============================================================================
// QUICK ACTION EMAIL FUNCTIONS - Send Forms, Surveys, and Status Updates
// ============================================================================

/**
 * Email the satisfaction survey link to a member
 * @param {string} memberId - Member ID to look up email
 */

/**
 * Email the contact info update form link to a member
 * @param {string} memberId - Member ID to look up email
 */

/**
 * Email the member dashboard/portal link to a member
 * @param {string} memberId - Member ID to look up email
 */

/**
 * Email grievance status update to the member
 * @param {string} grievanceId - Grievance ID to look up details
 */

/**
 * Helper: Get member data by Member ID
 * @private
 */

/**
 * Helper: Get organization name from Config
 * @private
 */

// ╔═══════════════════════════════════════════════════════════════════════════╗
// ║                                                                           ║
// ║   ██████╗ ██████╗  ██████╗ ████████╗███████╗ ██████╗████████╗███████╗██████╗  ║
// ║   ██╔══██╗██╔══██╗██╔═══██╗╚══██╔══╝██╔════╝██╔════╝╚══██╔══╝██╔════╝██╔══██╗ ║
// ║   ██████╔╝██████╔╝██║   ██║   ██║   █████╗  ██║        ██║   █████╗  ██║  ██║ ║
// ║   ██╔═══╝ ██╔══██╗██║   ██║   ██║   ██╔══╝  ██║        ██║   ██╔══╝  ██║  ██║ ║
// ║   ██║     ██║  ██║╚██████╔╝   ██║   ███████╗╚██████╗   ██║   ███████╗██████╔╝ ║
// ║   ╚═╝     ╚═╝  ╚═╝ ╚═════╝    ╚═╝   ╚══════╝ ╚═════╝   ╚═╝   ╚══════╝╚═════╝  ║
// ║                                                                           ║
// ║         ⚠️  DO NOT MODIFY THIS SECTION - PROTECTED CODE  ⚠️              ║
// ║                                                                           ║
// ╠═══════════════════════════════════════════════════════════════════════════╣
// ║  INTERACTIVE DASHBOARD TAB - Modal Popup with Tabbed Interface           ║
// ╠═══════════════════════════════════════════════════════════════════════════╣
// ║                                                                           ║
// ║  This code block is PROTECTED and should NOT be modified or removed.     ║
// ║                                                                           ║
// ║  Protected Functions:                                                     ║
// ║  • showInteractiveDashboardTab() - Opens the modal dialog                 ║
// ║  • getInteractiveDashboardHtml() - Returns the HTML/CSS/JS for the UI     ║
// ║  • getInteractiveOverviewData()  - Fetches overview statistics            ║
// ║  • getInteractiveMemberData()    - Fetches member list data               ║
// ║  • getInteractiveGrievanceData() - Fetches grievance list data            ║
// ║  • getInteractiveAnalyticsData() - Fetches analytics/charts data          ║
// ║                                                                           ║
// ║  Features:                                                                ║
// ║  • 4 Tabs: Overview, Members, Grievances, Analytics                       ║
// ║  • Live search and status filtering                                       ║
// ║  • Mobile-responsive design with touch targets                            ║
// ║  • Bar charts for status distribution and categories                      ║
// ║                                                                           ║
// ║  Menu Location: 👤 Dashboard > 🎯 Custom View                  ║
// ║                                                                           ║
// ║  Added: December 29, 2025 (commit c75c1cc)                                ║
// ║  Status: USER APPROVED - DO NOT CHANGE                                    ║
// ║                                                                           ║
// ╚═══════════════════════════════════════════════════════════════════════════╝

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Interactive dashboard is now consolidated into Steward Dashboard.
 */
function showInteractiveDashboardTab() {
  showStewardDashboard();
}

/**
 * Returns the HTML for the interactive dashboard with tabs
 */
function getInteractiveDashboardHtml() {
  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<base target="_top">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=5.0,user-scalable=yes">' +
    '<style>' +
    // CSS Reset and base styles
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;text-align:center}' +
    '.header h1{font-size:clamp(18px,4vw,24px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(11px,2.5vw,13px);opacity:0.9}' +

    // Tab navigation
    '.tabs{display:flex;background:white;border-bottom:2px solid #e0e0e0;position:sticky;top:0;z-index:100}' +
    '.tab{flex:1;padding:clamp(10px,2.5vw,14px);text-align:center;font-size:clamp(11px,2.5vw,13px);font-weight:600;color:#666;' +
    'border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.2s;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}' +
    '.tab:hover{background:#f8f9fa;color:#7C3AED}' +
    '.tab.active{color:#7C3AED;border-bottom-color:#7C3AED;background:#f8f4ff}' +
    '.tab-icon{display:block;font-size:16px;margin-bottom:2px}' +

    // Tab content
    '.tab-content{display:none;padding:15px;animation:fadeIn 0.3s}' +
    '.tab-content.active{display:block}' +
    '@keyframes fadeIn{from{opacity:0}to{opacity:1}}' +

    // Stats grid
    '.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(100px,1fr));gap:10px;margin-bottom:15px}' +
    '.stat-card{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;transition:transform 0.2s;cursor:pointer}' +
    '.stat-card:hover{transform:translateY(-2px)}' +
    '.stat-value{font-size:clamp(20px,4vw,28px);font-weight:bold;color:#7C3AED}' +
    '.stat-label{font-size:clamp(9px,1.8vw,11px);color:#666;text-transform:uppercase;margin-top:4px}' +
    '.stat-card.green .stat-value{color:#059669}' +
    '.stat-card.red .stat-value{color:#DC2626}' +
    '.stat-card.orange .stat-value{color:#F97316}' +

    // Data table
    '.data-table{width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08)}' +
    '.data-table th{background:#7C3AED;color:white;padding:12px;text-align:left;font-size:13px}' +
    '.data-table td{padding:12px;border-bottom:1px solid #eee;font-size:13px}' +
    '.data-table tr:hover{background:#f8f4ff}' +
    '.data-table tr:last-child td{border-bottom:none}' +

    // Status badges
    '.badge{display:inline-block;padding:4px 10px;border-radius:12px;font-size:11px;font-weight:bold}' +
    '.badge-open{background:#fee2e2;color:#dc2626}' +
    '.badge-pending{background:#fef3c7;color:#d97706}' +
    '.badge-closed{background:#d1fae5;color:#059669}' +
    '.badge-overdue{background:#7f1d1d;color:#fecaca}' +
    '.badge-steward{background:#ddd6fe;color:#7c3aed}' +

    // Action buttons
    '.action-btn{display:inline-flex;align-items:center;gap:6px;padding:8px 14px;border:none;border-radius:8px;' +
    'cursor:pointer;font-size:12px;font-weight:500;transition:all 0.2s;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}' +
    '.action-btn-primary{background:#7C3AED;color:white}' +
    '.action-btn-primary:hover{background:#5B21B6}' +
    '.action-btn-secondary{background:#f3f4f6;color:#374151}' +
    '.action-btn-secondary:hover{background:#e5e7eb}' +
    '.action-btn-danger{background:#dc2626;color:white}' +
    '.action-btn-danger:hover{background:#b91c1c}' +
    '.action-btn.active{background:#7C3AED;color:white}' +

    // List items - clickable
    '.list-container{display:flex;flex-direction:column;gap:10px}' +
    '.list-item{background:white;padding:15px;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,0.06);cursor:pointer;transition:all 0.2s}' +
    '.list-item:hover{box-shadow:0 4px 12px rgba(0,0,0,0.12);transform:translateY(-1px)}' +
    '.list-item-header{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;flex-wrap:wrap}' +
    '.list-item-main{flex:1;min-width:180px}' +
    '.list-item-title{font-weight:600;color:#1f2937;margin-bottom:3px}' +
    '.list-item-subtitle{font-size:12px;color:#666}' +
    '.list-item-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee;font-size:12px;color:#374151}' +
    '.list-item.expanded .list-item-details{display:block}' +
    '.detail-row{display:flex;gap:8px;margin-bottom:6px}' +
    '.detail-label{color:#666;min-width:90px}' +
    '.detail-value{color:#1f2937;font-weight:500}' +
    '.detail-actions{margin-top:10px;display:flex;gap:8px;flex-wrap:wrap}' +

    // Filter dropdowns
    '.filter-bar{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center}' +
    '.filter-select{padding:8px 12px;border:2px solid #e5e7eb;border-radius:6px;font-size:12px;background:white;cursor:pointer}' +
    '.filter-select:focus{outline:none;border-color:#7C3AED}' +

    // Search input
    '.sr-only{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;border:0}' +
    '.search-container{position:relative;margin-bottom:12px}' +
    '.search-input{width:100%;padding:10px 10px 10px 36px;border:2px solid #e5e7eb;border-radius:8px;font-size:13px;transition:border-color 0.2s}' +
    '.search-input:focus{outline:none;border-color:#7C3AED}' +
    '.search-icon{position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:14px;color:#9ca3af}' +

    // Resource links
    '.resource-links{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);margin-top:15px}' +
    '.resource-links h3{font-size:14px;color:#1f2937;margin-bottom:12px}' +
    '.link-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px}' +
    '.resource-link{display:flex;align-items:center;gap:8px;padding:10px;background:#f8f4ff;border-radius:8px;text-decoration:none;color:#7C3AED;font-size:12px;font-weight:500;transition:all 0.2s}' +
    '.resource-link:hover{background:#7C3AED;color:white}' +

    // Charts section
    '.chart-container{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);margin-bottom:12px}' +
    '.chart-title{font-weight:600;color:#1f2937;margin-bottom:12px;font-size:13px}' +
    '.bar-chart{display:flex;flex-direction:column;gap:8px}' +
    '.bar-row{display:flex;align-items:center;gap:8px}' +
    '.bar-label{width:90px;font-size:11px;color:#666;text-align:right}' +
    '.bar-container{flex:1;background:#e5e7eb;border-radius:4px;height:20px;overflow:hidden}' +
    '.bar-fill{height:100%;border-radius:4px;transition:width 0.5s}' +
    '.bar-value{width:40px;font-size:11px;font-weight:600;color:#374151}' +

    // Empty state
    '.empty-state{text-align:center;padding:30px;color:#9ca3af}' +
    '.empty-state-icon{font-size:40px;margin-bottom:8px}' +

    // Loading
    '.loading{text-align:center;padding:30px;color:#666}' +
    '.spinner{display:inline-block;width:20px;height:20px;border:3px solid #e5e7eb;border-top-color:#7C3AED;border-radius:50%;animation:spin 1s linear infinite}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +

    // Error state
    '.error-state{text-align:center;padding:30px;color:#dc2626;background:#fef2f2;border-radius:8px;margin:10px;border:1px solid #fecaca}' +
    '.error-state::before{content:"⚠️ ";font-size:20px}' +
    '.loading-state{text-align:center;padding:40px;color:#6b7280}' +
    '.loading-spinner{display:inline-block;width:24px;height:24px;border:3px solid #e5e7eb;border-top-color:#7c3aed;border-radius:50%;animation:spin 1s linear infinite;margin-bottom:10px}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.debug-info{font-size:10px;color:#9ca3af;margin-top:5px;font-family:monospace}' +

    // Sankey Diagram
    '.sankey-container{position:relative;padding:15px 0}' +
    '.sankey-nodes{display:flex;justify-content:space-between;position:relative;z-index:2}' +
    '.sankey-column{display:flex;flex-direction:column;gap:6px;align-items:center}' +
    '.sankey-node{padding:8px 12px;border-radius:6px;color:white;font-weight:600;font-size:11px;text-align:center;min-width:70px;box-shadow:0 2px 4px rgba(0,0,0,0.2)}' +
    '.sankey-node.source{background:linear-gradient(135deg,#7C3AED,#9333EA)}' +
    '.sankey-node.status-open{background:linear-gradient(135deg,#dc2626,#ef4444)}' +
    '.sankey-node.status-pending{background:linear-gradient(135deg,#f97316,#fb923c)}' +
    '.sankey-node.status-closed{background:linear-gradient(135deg,#059669,#10b981)}' +
    '.sankey-node.resolution{background:linear-gradient(135deg,#1d4ed8,#3b82f6)}' +
    '.sankey-flows{position:absolute;top:0;left:0;right:0;bottom:0;z-index:1}' +
    '.sankey-flow{position:absolute;height:4px;border-radius:2px;opacity:0.6;transition:opacity 0.2s}' +
    '.sankey-flow:hover{opacity:1}' +
    '.sankey-label{font-size:10px;color:#666;margin-top:4px}' +
    '.sankey-legend{display:flex;flex-wrap:wrap;gap:10px;justify-content:center;margin-top:12px;padding-top:12px;border-top:1px solid #e5e7eb}' +
    '.sankey-legend-item{display:flex;align-items:center;gap:4px;font-size:10px;color:#666}' +
    '.sankey-legend-color{width:10px;height:10px;border-radius:2px}' +

    // Responsive
    '@media (max-width:600px){' +
    '  .stats-grid{grid-template-columns:repeat(2,1fr)}' +
    '  .list-item-header{flex-direction:column;align-items:flex-start}' +
    '  .tab-icon{font-size:14px}' +
    '  .filter-bar{flex-direction:column}' +
    '  .filter-select{width:100%}' +
    '}' +

    '</style>' +
    '</head><body>' +

    // Header
    '<div class="header">' +
    '<h1>📊 Dashboard</h1>' +
    '<div class="subtitle">Real-time union data at your fingertips</div>' +
    '</div>' +

    // Tab Navigation (6 tabs now - including My Cases)
    '<div class="tabs" role="tablist" aria-label="Dashboard navigation">' +
    '<button class="tab active" role="tab" aria-selected="true" aria-controls="content-overview" onclick="switchTab(\'overview\',this)" id="tab-overview"><span class="tab-icon">📊</span>Overview</button>' +
    '<button class="tab" role="tab" aria-selected="false" aria-controls="content-mycases" onclick="switchTab(\'mycases\',this)" id="tab-mycases"><span class="tab-icon">👤</span>My Cases</button>' +
    '<button class="tab" role="tab" aria-selected="false" aria-controls="content-members" onclick="switchTab(\'members\',this)" id="tab-members"><span class="tab-icon">👥</span>Members</button>' +
    '<button class="tab" role="tab" aria-selected="false" aria-controls="content-grievances" onclick="switchTab(\'grievances\',this)" id="tab-grievances"><span class="tab-icon">📋</span>Grievances</button>' +
    '<button class="tab" role="tab" aria-selected="false" aria-controls="content-analytics" onclick="switchTab(\'analytics\',this)" id="tab-analytics"><span class="tab-icon">📈</span>Analytics</button>' +
    '<button class="tab" role="tab" aria-selected="false" aria-controls="content-resources" onclick="switchTab(\'resources\',this)" id="tab-resources"><span class="tab-icon">🔗</span>Links</button>' +
    '</div>' +

    // Overview Tab
    '<div class="tab-content active" id="content-overview" role="tabpanel" aria-labelledby="tab-overview">' +
    '<div class="stats-grid" id="overview-stats"><div class="loading"><div class="spinner"></div><p>Loading stats...</p></div></div>' +
    '<div id="overview-actions" style="margin-top:12px;display:flex;flex-wrap:wrap;gap:8px"></div>' +
    '<div id="overview-overdue" style="margin-top:15px"></div>' +
    '</div>' +

    // My Cases Tab - Shows steward's assigned grievances
    '<div class="tab-content" id="content-mycases" role="tabpanel" aria-labelledby="tab-mycases">' +
    '<div class="section-card" style="background:linear-gradient(135deg,#f0f4ff,#e8f0fe);border-left:4px solid #7C3AED;margin-bottom:15px">' +
    '<div style="display:flex;align-items:center;gap:10px"><span style="font-size:24px">👤</span><div><strong>My Assigned Cases</strong><div style="font-size:12px;color:#666">Grievances where you are the assigned steward</div></div></div>' +
    '</div>' +
    '<div class="filter-bar" id="mycases-filter-bar">' +
    '<button class="action-btn action-btn-primary active" data-filter="all" onclick="filterMyCasesStatus(\'all\',this)">All My Cases</button>' +
    '<button class="action-btn action-btn-secondary" data-filter="Open" onclick="filterMyCasesStatus(\'Open\',this)">Open</button>' +
    '<button class="action-btn action-btn-secondary" data-filter="Pending Info" onclick="filterMyCasesStatus(\'Pending Info\',this)">Pending</button>' +
    '<button class="action-btn action-btn-danger" data-filter="Overdue" onclick="filterMyCasesStatus(\'Overdue\',this)">⚠️ Overdue</button>' +
    '</div>' +
    '<div id="mycases-stats" style="margin-bottom:15px"></div>' +
    '<div class="list-container" id="mycases-list"><div class="loading"><div class="spinner"></div><p>Loading your cases...</p></div></div>' +
    '</div>' +

    // Members Tab
    '<div class="tab-content" id="content-members" role="tabpanel" aria-labelledby="tab-members">' +
    '<div class="search-container"><span class="search-icon" aria-hidden="true">🔍</span><label for="member-search" class="sr-only">Search members</label><input type="text" class="search-input" id="member-search" placeholder="Search by name, ID, title, location..." oninput="filterMembers()" aria-label="Search members by name, ID, title, or location"></div>' +
    '<div class="filter-bar" id="member-filters"></div>' +
    '<div style="margin-bottom:12px"><button class="action-btn action-btn-primary" onclick="showAddMemberForm()">➕ Add New Member</button></div>' +
    '<div class="list-container" id="members-list"><div class="loading"><div class="spinner"></div><p>Loading members...</p></div></div>' +
    // Add Member Form Modal (hidden initially)
    '<div id="member-form-modal" style="display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:1000;overflow-y:auto;padding:20px">' +
    '<div style="background:white;max-width:500px;margin:20px auto;border-radius:12px;padding:20px;box-shadow:0 10px 40px rgba(0,0,0,0.2)">' +
    '<h3 id="member-form-title" style="margin:0 0 15px;color:#7C3AED">➕ Add New Member</h3>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">First Name *</label><input type="text" id="form-firstName" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" placeholder="Enter first name"></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Last Name *</label><input type="text" id="form-lastName" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" placeholder="Enter last name"></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Job Title</label><select id="form-jobTitle" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px"><option value="">Select job title...</option></select></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Email</label><input type="email" id="form-email" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" placeholder="Enter email address"></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Phone</label><input type="tel" id="form-phone" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" placeholder="Enter phone number"></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Work Location</label><select id="form-location" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px"><option value="">Select location...</option></select></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Unit</label><select id="form-unit" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px"><option value="">Select unit...</option></select></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Office Days</label><select id="form-officeDays" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" multiple size="3"></select><small style="color:#999;font-size:10px">Hold Ctrl/Cmd to select multiple days</small></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Supervisor</label><select id="form-supervisor" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px"><option value="">Select supervisor...</option></select></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Manager</label><select id="form-manager" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px"><option value="">Select manager...</option></select></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Is Steward?</label><select id="form-isSteward" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px"><option value="No">No</option><option value="Yes">Yes</option></select></div>' +
    '<input type="hidden" id="form-memberId" value="">' +
    '<input type="hidden" id="form-mode" value="add">' +
    '<div style="display:flex;gap:10px;margin-top:20px">' +
    '<button class="action-btn action-btn-primary" style="flex:1" onclick="saveMemberForm()">💾 Save Member</button>' +
    '<button class="action-btn action-btn-secondary" style="flex:1" onclick="closeMemberForm()">Cancel</button>' +
    '</div>' +
    '</div></div>' +
    '</div>' +

    // Grievances Tab
    '<div class="tab-content" id="content-grievances" role="tabpanel" aria-labelledby="tab-grievances">' +
    '<div class="search-container"><span class="search-icon" aria-hidden="true">🔍</span><label for="grievance-search" class="sr-only">Search grievances</label><input type="text" class="search-input" id="grievance-search" placeholder="Search by ID, member name, issue..." oninput="filterGrievances()" aria-label="Search grievances by ID, member name, or issue"></div>' +
    '<div class="filter-bar" id="grievance-filter-bar">' +
    '<button class="action-btn action-btn-primary active" data-filter="all" onclick="filterGrievanceStatus(\'all\',this)">All</button>' +
    '<button class="action-btn action-btn-secondary" data-filter="Open" onclick="filterGrievanceStatus(\'Open\',this)">Open</button>' +
    '<button class="action-btn action-btn-secondary" data-filter="Pending Info" onclick="filterGrievanceStatus(\'Pending Info\',this)">Pending</button>' +
    '<button class="action-btn action-btn-danger" data-filter="Overdue" onclick="filterGrievanceStatus(\'Overdue\',this)">⚠️ Overdue</button>' +
    '<button class="action-btn action-btn-secondary" data-filter="Closed" onclick="filterGrievanceStatus(\'Closed\',this)">Closed</button>' +
    '</div>' +
    '<div class="list-container" id="grievances-list"><div class="loading"><div class="spinner"></div><p>Loading grievances...</p></div></div>' +
    '</div>' +

    // Analytics Tab
    '<div class="tab-content" id="content-analytics" role="tabpanel" aria-labelledby="tab-analytics">' +
    '<div id="analytics-charts"><div class="loading"><div class="spinner"></div><p>Loading analytics...</p></div></div>' +
    '</div>' +

    // Resources Tab
    '<div class="tab-content" id="content-resources" role="tabpanel" aria-labelledby="tab-resources">' +
    '<div id="resources-content"><div class="loading"><div class="spinner"></div><p>Loading links...</p></div></div>' +
    '</div>' +

    // JavaScript
    '<script>' +
    // XSS Prevention - escape HTML special characters
    getClientSideEscapeHtml() +
    'var allMembers=[];var allGrievances=[];var myCases=[];var currentGrievanceFilter="all";var currentMyCasesFilter="all";var memberFilters={location:"all",unit:"all",officeDays:"all"};var resourceLinks={};' +

    // Debug mode and error handler wrapper
    'var DEBUG_MODE=false;' +
    'function log(msg,data){if(DEBUG_MODE){console.log("[Dashboard] "+msg,data||"")}}' +
    'function logError(msg,e){console.error("[Dashboard Error] "+msg,e);if(DEBUG_MODE)alert("Debug: "+msg+"\\n"+escapeHtml(e.message))}' +
    'function safeRun(fn,fallback){try{fn()}catch(e){console.error("[Dashboard]",e);if(fallback)fallback(e)}}' +
    'function showLoading(elementId,msg){var el=document.getElementById(elementId);if(el)el.innerHTML="<div class=\\"loading-state\\"><div class=\\"loading-spinner\\"></div><div>"+escapeHtml(msg||"Loading...")+"</div></div>"}' +
    'function safeUrl(url){if(!url)return "#";var s=String(url).trim();return(/^https?:\\/\\//i.test(s))?escapeHtml(s):"#"}' +

    // Tab switching with error handling
    'function switchTab(tabName,btn){' +
    '  safeRun(function(){' +
    '    document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active");t.setAttribute("aria-selected","false")});' +
    '    document.querySelectorAll(".tab-content").forEach(function(c){c.classList.remove("active")});' +
    '    btn.classList.add("active");' +
    '    btn.setAttribute("aria-selected","true");' +
    '    document.getElementById("content-"+tabName).classList.add("active");' +
    '    if(tabName==="mycases"&&myCases.length===0)loadMyCases();' +
    '    if(tabName==="members"&&allMembers.length===0)loadMembers();' +
    '    if(tabName==="grievances"&&allGrievances.length===0)loadGrievances();' +
    '    if(tabName==="analytics")loadAnalytics();' +
    '    if(tabName==="resources")loadResources();' +
    '  });' +
    '}' +

    // Keyboard navigation for tabs
    'document.addEventListener("keydown",function(e){' +
    '  var activeTab=document.querySelector(".tab.active");' +
    '  if(!activeTab)return;' +
    '  var tabs=Array.prototype.slice.call(document.querySelectorAll(".tab"));' +
    '  var idx=tabs.indexOf(activeTab);' +
    '  if(idx===-1)return;' +
    '  if(e.key==="ArrowRight"||e.key==="ArrowDown"){' +
    '    e.preventDefault();' +
    '    var next=tabs[(idx+1)%tabs.length];' +
    '    next.click();' +
    '    next.focus();' +
    '  }else if(e.key==="ArrowLeft"||e.key==="ArrowUp"){' +
    '    e.preventDefault();' +
    '    var prev=tabs[(idx-1+tabs.length)%tabs.length];' +
    '    prev.click();' +
    '    prev.focus();' +
    '  }' +
    '});' +

    // Load overview data with error handling
    'function loadOverview(){' +
    '  log("Loading overview data...");' +
    '  showLoading("overview-stats","Loading dashboard stats...");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){' +
    '      log("Overview data received:",data);' +
    '      safeRun(function(){renderOverview(data)},function(e){' +
    '        document.getElementById("overview-stats").innerHTML="<div class=\\"error-state\\">Error rendering stats<div class=\\"debug-info\\">"+escapeHtml(e.message)+"</div></div>"' +
    '      })' +
    '    })'  +
    '    .withFailureHandler(function(e){' +
    '      logError("Failed to load overview",e);' +
    '      document.getElementById("overview-stats").innerHTML="<div class=\\"error-state\\">Failed to load stats<div class=\\"debug-info\\">"+escapeHtml(e.message)+"<br>Check: Admin → Modal Diagnostics</div></div>"' +
    '    })' +
    '    .getInteractiveOverviewData();' +
    '}' +

    // Render overview with overdue section and location breakdown
    'function renderOverview(data){' +
    '  var html="";' +
    '  var colors=["#7C3AED","#059669","#1a73e8","#F97316","#DC2626","#8B5CF6","#10B981","#3B82F6"];' +
    '  html+="<div class=\\"stat-card\\" onclick=\\"switchTab(\'members\',document.getElementById(\'tab-members\'))\\"><div class=\\"stat-value\\">"+data.totalMembers+"</div><div class=\\"stat-label\\">Total Members</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.activeStewards+"</div><div class=\\"stat-label\\">Stewards</div></div>";' +
    '  html+="<div class=\\"stat-card\\" onclick=\\"switchTab(\'grievances\',document.getElementById(\'tab-grievances\'))\\"><div class=\\"stat-value\\">"+data.totalGrievances+"</div><div class=\\"stat-label\\">Total Grievances</div></div>";' +
    '  html+="<div class=\\"stat-card red\\" onclick=\\"showOpenCases()\\"><div class=\\"stat-value\\">"+data.openGrievances+"</div><div class=\\"stat-label\\">Open Cases</div></div>";' +
    '  html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+data.pendingInfo+"</div><div class=\\"stat-label\\">Pending Info</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.winRate+"</div><div class=\\"stat-label\\">Win Rate</div></div>";' +
    '  document.getElementById("overview-stats").innerHTML=html;' +
    '  var actions="";' +
    '  actions+="<button class=\\"action-btn action-btn-primary\\" onclick=\\"google.script.run.showMobileUnifiedSearch()\\">🔍 Search</button>";' +
    '  actions+="<button class=\\"action-btn action-btn-secondary\\" onclick=\\"google.script.run.showMobileGrievanceList()\\">📋 All Grievances</button>";' +
    '  actions+="<button class=\\"action-btn action-btn-secondary\\" onclick=\\"google.script.run.showMyAssignedGrievances()\\">👤 My Cases</button>";' +
    '  actions+="<button class=\\"action-btn action-btn-secondary\\" onclick=\\"location.reload()\\">🔄 Refresh</button>";' +
    '  document.getElementById("overview-actions").innerHTML=actions;' +
    // Location breakdown with bubble chart
    '  if(data.byLocation&&data.byLocation.length>0){' +
    '    var locHtml="<div class=\\"chart-container\\" style=\\"margin-top:15px\\"><div class=\\"chart-title\\">📍 Members by Location</div>";' +
    '    locHtml+="<div style=\\"display:flex;flex-wrap:wrap;gap:10px;justify-content:center;padding:10px\\">";' +
    '    var maxLoc=Math.max.apply(null,data.byLocation.map(function(l){return l.count}))||1;' +
    '    var totalM=data.totalMembers||1;' +
    '    data.byLocation.forEach(function(loc,idx){' +
    '      var pct=Math.round(loc.count/totalM*100);' +
    '      var size=Math.max(55,Math.min(100,50+(loc.count/maxLoc*50)));' +
    '      var clr=colors[idx%colors.length];' +
    '      locHtml+="<div style=\\"text-align:center;cursor:pointer\\" onclick=\\"switchTab(\'analytics\',document.getElementById(\'tab-analytics\'))\\"><div style=\\"width:"+size+"px;height:"+size+"px;border-radius:50%;background:"+clr+";display:flex;flex-direction:column;align-items:center;justify-content:center;color:white;font-weight:bold;margin:0 auto;box-shadow:0 3px 10px rgba(0,0,0,0.15);transition:transform 0.2s\\" onmouseover=\\"this.style.transform=\'scale(1.05)\'\\" onmouseout=\\"this.style.transform=\'scale(1)\'\\"><span style=\\"font-size:"+(size/3)+"px\\">"+loc.count+"</span><span style=\\"font-size:9px;opacity:0.9\\">"+pct+"%</span></div><div style=\\"font-size:10px;color:#666;margin-top:5px;max-width:80px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap\\">"+escapeHtml(loc.name)+"</div></div>";' +
    '    });' +
    '    locHtml+="</div></div>";' +
    '    document.getElementById("overview-actions").insertAdjacentHTML("afterend",locHtml);' +
    '  }' +
    '  loadOverduePreview();' +
    '}' +

    // Show open cases - switch to grievances tab with Open filter
    'function showOpenCases(){switchTab("grievances",document.getElementById("tab-grievances"));setTimeout(function(){filterGrievanceStatus("Open",document.querySelector("[data-filter=\\"Open\\"]"))},300)}' +

    // Load overdue preview on overview
    'function loadOverduePreview(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    var overdue=data.filter(function(g){return g.isOverdue});' +
    '    if(overdue.length===0){document.getElementById("overview-overdue").innerHTML="";return}' +
    '    var html="<div class=\\"chart-container\\" style=\\"border-left:4px solid #dc2626\\"><div class=\\"chart-title\\">⚠️ Overdue Cases ("+overdue.length+")</div>";' +
    '    html+="<div class=\\"list-container\\">";' +
    '    overdue.slice(0,3).forEach(function(g){html+="<div class=\\"list-item\\" onclick=\\"showGrievanceDetail(\'"+escapeHtml(g.id)+"\')\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+escapeHtml(g.id)+" - "+escapeHtml(g.memberName)+"</div><div class=\\"list-item-subtitle\\">"+escapeHtml(g.issueType)+" • "+escapeHtml(g.currentStep)+"</div></div><span class=\\"badge badge-overdue\\">Overdue</span></div>"});' +
    '    if(overdue.length>3)html+="<button class=\\"action-btn action-btn-danger\\" style=\\"width:100%;margin-top:8px\\" onclick=\\"switchTab(\'grievances\',document.getElementById(\'tab-grievances\'));setTimeout(function(){filterGrievanceStatus(\'Overdue\',document.querySelector(\'[data-filter=Overdue]\'))},300)\\">View All "+overdue.length+" Overdue Cases</button>";' +
    '    html+="</div></div>";' +
    '    document.getElementById("overview-overdue").innerHTML=html;' +
    '  }).getInteractiveGrievanceData();' +
    '}' +

    // Load my cases (steward's assigned grievances)
    'function loadMyCases(){' +
    '  log("Loading my cases...");' +
    '  showLoading("mycases-list","Loading your assigned cases...");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){log("My cases received:",data?data.length:0);myCases=data||[];renderMyCases(myCases);renderMyCasesStats(data)})'  +
    '    .withFailureHandler(function(e){logError("Failed to load my cases",e);document.getElementById("mycases-list").innerHTML="<div class=\\"error-state\\">Failed to load your cases<div class=\\"debug-info\\">"+escapeHtml(e.message)+"</div></div>"})' +
    '    .getMyStewardCases();' +
    '}' +

    // Render my cases stats
    'function renderMyCasesStats(data){' +
    '  var total=data.length;' +
    '  var open=data.filter(function(g){return g.status==="Open"}).length;' +
    '  var pending=data.filter(function(g){return g.status==="Pending Info"}).length;' +
    '  var overdue=data.filter(function(g){return g.isOverdue}).length;' +
    '  var html="<div class=\\"stats-grid\\">";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+total+"</div><div class=\\"stat-label\\">Total Assigned</div></div>";' +
    '  html+="<div class=\\"stat-card red\\"><div class=\\"stat-value\\">"+open+"</div><div class=\\"stat-label\\">Open</div></div>";' +
    '  html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+pending+"</div><div class=\\"stat-label\\">Pending Info</div></div>";' +
    '  if(overdue>0)html+="<div class=\\"stat-card\\" style=\\"border:2px solid #dc2626\\"><div class=\\"stat-value\\" style=\\"color:#dc2626\\">"+overdue+"</div><div class=\\"stat-label\\">⚠️ Overdue</div></div>";' +
    '  html+="</div>";' +
    '  document.getElementById("mycases-stats").innerHTML=html;' +
    '}' +

    // Render my cases list
    'function renderMyCases(data){' +
    '  var c=document.getElementById("mycases-list");' +
    '  if(!data||data.length===0){c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">👤</div><p>No cases assigned to you</p><p style=\\"font-size:12px;color:#999;margin-top:8px\\">Cases where you are listed as the steward will appear here</p></div>";return}' +
    '  c.innerHTML=data.map(function(g,i){' +
    '    var badgeClass=g.isOverdue?"badge-overdue":(g.status==="Open"?"badge-open":(g.status==="Pending Info"?"badge-pending":"badge-closed"));' +
    '    var statusText=g.isOverdue?"Overdue":g.status;' +
    '    var priorityBorder=g.isOverdue?"border-left:4px solid #dc2626;":"";' +
    '    return "<div class=\\"list-item\\" style=\\""+priorityBorder+"\\" onclick=\\"toggleMyCaseDetail(this)\\">' +
    '      <div class=\\"list-item-header\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+escapeHtml(g.id)+" - "+escapeHtml(g.memberName)+"</div><div class=\\"list-item-subtitle\\">"+escapeHtml(g.issueType)+" • "+escapeHtml(g.currentStep)+"</div></div><div><span class=\\"badge "+badgeClass+"\\">"+escapeHtml(statusText)+"</span></div></div>' +
    '      <div class=\\"list-item-details\\">' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📅 Filed:</span><span class=\\"detail-value\\">"+escapeHtml(g.filedDate)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">⏰ Next Due:</span><span class=\\"detail-value\\">"+escapeHtml(g.nextActionDue)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">⏱️ Days Open:</span><span class=\\"detail-value\\">"+escapeHtml(g.daysOpen)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📍 Location:</span><span class=\\"detail-value\\">"+escapeHtml(g.location)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📜 Articles:</span><span class=\\"detail-value\\">"+escapeHtml(g.articles)+"</span></div>' +
    '        <div class=\\"detail-actions\\">' +
    '          <button class=\\"action-btn action-btn-primary\\" data-gid=\\""+escapeHtml(g.id)+"\\" onclick=\\"event.stopPropagation();google.script.run.showGrievanceQuickActions(this.dataset.gid)\\">⚡ Quick Actions</button>' +
    '          <button class=\\"action-btn action-btn-secondary\\" data-gid=\\""+escapeHtml(g.id)+"\\" onclick=\\"event.stopPropagation();google.script.run.navigateToGrievanceInSheet(this.dataset.gid)\\">📄 View in Sheet</button>' +
    '        </div>' +
    '      </div>' +
    '    </div>"' +
    '  }).join("");' +
    '}' +

    // Toggle my case detail
    'function toggleMyCaseDetail(el){el.classList.toggle("expanded")}' +

    // Filter my cases by status
    'function filterMyCasesStatus(status,btn){' +
    '  currentMyCasesFilter=status;' +
    '  document.querySelectorAll("#mycases-filter-bar .action-btn").forEach(function(b){' +
    '    b.classList.remove("active","action-btn-primary");' +
    '    if(b.dataset.filter!=="Overdue")b.classList.add("action-btn-secondary");' +
    '  });' +
    '  if(btn){btn.classList.add("active");if(status!=="Overdue")btn.classList.add("action-btn-primary");btn.classList.remove("action-btn-secondary")}' +
    '  var filtered=myCases;' +
    '  if(status==="Overdue"){filtered=myCases.filter(function(g){return g.isOverdue})}' +
    '  else if(status!=="all"){filtered=myCases.filter(function(g){return g.status===status})}' +
    '  renderMyCases(filtered);' +
    '}' +

    // Config dropdown values (populated by loadConfigDropdowns)
    'var configDropdowns={};' +

    // Load Config dropdown values for form selects
    'function loadConfigDropdowns(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){configDropdowns=data||{};populateFormDropdowns()})' +
    '    .withFailureHandler(function(e){logError("Failed to load config dropdowns",e)})' +
    '    .getConfigDropdownValues();' +
    '}' +

    // Load members with filters
    'function loadMembers(){' +
    '  log("Loading members...");' +
    '  showLoading("members-list","Loading member directory...");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){log("Members received:",data?data.length:0);allMembers=data||[];renderMembers(allMembers);loadMemberFilters();loadConfigDropdowns()})'  +
    '    .withFailureHandler(function(e){logError("Failed to load members",e);document.getElementById("members-list").innerHTML="<div class=\\"error-state\\">Failed to load members<div class=\\"debug-info\\">"+escapeHtml(e.message)+"</div></div>"})' +
    '    .getInteractiveMemberData();' +
    '}' +

    // Load member filter dropdowns
    'function loadMemberFilters(){' +
    '  var locations={};var units={};var officeDays={};' +
    '  allMembers.forEach(function(m){' +
    '    if(m.location&&m.location!=="N/A")locations[m.location]=1;' +
    '    if(m.unit&&m.unit!=="N/A")units[m.unit]=1;' +
    '    if(m.officeDays&&m.officeDays!=="N/A"){' +
    '      m.officeDays.split(",").forEach(function(d){var day=d.trim();if(day)officeDays[day]=1});' +
    '    }' +
    '  });' +
    '  var html="<select class=\\"filter-select\\" id=\\"filter-location\\" onchange=\\"memberFilters.location=this.value;filterMembers()\\"><option value=\\"all\\">All Locations</option>";' +
    '  Object.keys(locations).sort().forEach(function(l){html+="<option value=\\""+escapeHtml(l)+"\\">"+escapeHtml(l)+"</option>"});' +
    '  html+="</select><select class=\\"filter-select\\" id=\\"filter-unit\\" onchange=\\"memberFilters.unit=this.value;filterMembers()\\"><option value=\\"all\\">All Units</option>";' +
    '  Object.keys(units).sort().forEach(function(u){html+="<option value=\\""+escapeHtml(u)+"\\">"+escapeHtml(u)+"</option>"});' +
    '  html+="</select><select class=\\"filter-select\\" id=\\"filter-officeDays\\" onchange=\\"memberFilters.officeDays=this.value;filterMembers()\\"><option value=\\"all\\">All Office Days</option>";' +
    '  Object.keys(officeDays).sort(function(a,b){var days=[\"Monday\",\"Tuesday\",\"Wednesday\",\"Thursday\",\"Friday\",\"Saturday\",\"Sunday\"];return days.indexOf(a)-days.indexOf(b)}).forEach(function(d){html+="<option value=\\""+escapeHtml(d)+"\\">"+escapeHtml(d)+"</option>"});' +
    '  html+="</select><button class=\\"action-btn action-btn-secondary\\" onclick=\\"resetMemberFilters()\\">Reset</button>";' +
    '  document.getElementById("member-filters").innerHTML=html;' +
    '  populateFormDropdowns(locations,units);' +
    '}' +

    // Reset member filters
    'function resetMemberFilters(){memberFilters={location:"all",unit:"all",officeDays:"all"};document.getElementById("member-search").value="";document.getElementById("filter-location").value="all";document.getElementById("filter-unit").value="all";document.getElementById("filter-officeDays").value="all";renderMembers(allMembers)}' +

    // Render members with clickable details
    'function renderMembers(data){' +
    '  var c=document.getElementById("members-list");' +
    '  if(!data||data.length===0){c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">👥</div><p>No members found</p></div>";return}' +
    '  c.innerHTML=data.slice(0,50).map(function(m,i){' +
    '    var badge=m.isSteward?"<span class=\\"badge badge-steward\\">Steward</span>":"";' +
    '    if(m.hasOpenGrievance)badge+="<span class=\\"badge badge-open\\" style=\\"margin-left:4px\\">Has Case</span>";' +
    '    return "<div class=\\"list-item\\" onclick=\\"toggleMemberDetail(this,"+i+")\\">' +
    '      <div class=\\"list-item-header\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+escapeHtml(m.name)+"</div><div class=\\"list-item-subtitle\\">"+escapeHtml(m.id)+" • "+escapeHtml(m.title)+"</div></div><div>"+badge+"</div></div>' +
    '      <div class=\\"list-item-details\\">' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📍 Location:</span><span class=\\"detail-value\\">"+escapeHtml(m.location)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">🏢 Unit:</span><span class=\\"detail-value\\">"+escapeHtml(m.unit)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📧 Email:</span><span class=\\"detail-value\\">"+escapeHtml(m.email||"N/A")+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📱 Phone:</span><span class=\\"detail-value\\">"+escapeHtml(m.phone||"N/A")+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📅 Office Days:</span><span class=\\"detail-value\\">"+escapeHtml(m.officeDays)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">👤 Supervisor:</span><span class=\\"detail-value\\">"+escapeHtml(m.supervisor)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">🛡️ Steward:</span><span class=\\"detail-value\\">"+escapeHtml(m.assignedSteward)+"</span></div>' +
    '        <div class=\\"detail-actions\\">' +
    '          <button class=\\"action-btn action-btn-primary\\" onclick=\\"event.stopPropagation();showEditMemberForm("+i+")\\">✏️ Edit Member</button>' +
    '          <button class=\\"action-btn action-btn-secondary\\" onclick=\\"event.stopPropagation();google.script.run.navigateToMemberInSheet(\'"+escapeHtml(m.id).replace(/\'/g,"")+"\')\\">📄 View in Sheet</button>' +
    '        </div>' +
    '      </div>' +
    '    </div>"' +
    '  }).join("");' +
    '  if(data.length>50)c.innerHTML+="<div class=\\"empty-state\\"><p>Showing 50 of "+data.length+" members. Use search to find specific members.</p></div>";' +
    '}' +

    // Toggle member detail
    'function toggleMemberDetail(el,idx){el.classList.toggle("expanded")}' +

    // Filter members with all criteria
    'function filterMembers(){' +
    '  var query=(document.getElementById("member-search").value||"").toLowerCase();' +
    '  var filtered=allMembers.filter(function(m){' +
    '    if(memberFilters.location!=="all"&&m.location!==memberFilters.location)return false;' +
    '    if(memberFilters.unit!=="all"&&m.unit!==memberFilters.unit)return false;' +
    '    if(memberFilters.officeDays!=="all"&&m.officeDays&&m.officeDays.indexOf(memberFilters.officeDays)<0)return false;' +
    '    if(query&&query.length>=2){' +
    '      return m.name.toLowerCase().indexOf(query)>=0||' +
    '             m.id.toLowerCase().indexOf(query)>=0||' +
    '             m.title.toLowerCase().indexOf(query)>=0||' +
    '             m.location.toLowerCase().indexOf(query)>=0||' +
    '             (m.email||"").toLowerCase().indexOf(query)>=0;' +
    '    }' +
    '    return true;' +
    '  });' +
    '  renderMembers(filtered);' +
    '}' +

    // Populate form dropdowns from Config sheet values
    'function populateFormDropdowns(){' +
    '  function fillSelect(id,placeholder,values){' +
    '    var sel=document.getElementById(id);if(!sel)return;' +
    '    sel.innerHTML="<option value=\\"\\">"+escapeHtml(placeholder)+"</option>";' +
    '    (values||[]).forEach(function(v){sel.innerHTML+="<option value=\\""+escapeHtml(v)+"\\">"+escapeHtml(v)+"</option>"});' +
    '  }' +
    '  function fillMultiSelect(id,values){' +
    '    var sel=document.getElementById(id);if(!sel)return;' +
    '    sel.innerHTML="";' +
    '    (values||[]).forEach(function(v){sel.innerHTML+="<option value=\\""+escapeHtml(v)+"\\">"+escapeHtml(v)+"</option>"});' +
    '  }' +
    '  fillSelect("form-jobTitle","Select job title...",configDropdowns.jobTitles);' +
    '  fillSelect("form-location","Select location...",configDropdowns.locations);' +
    '  fillSelect("form-unit","Select unit...",configDropdowns.units);' +
    '  fillMultiSelect("form-officeDays",configDropdowns.officeDays);' +
    '  fillSelect("form-supervisor","Select supervisor...",configDropdowns.supervisors);' +
    '  fillSelect("form-manager","Select manager...",configDropdowns.managers);' +
    '}' +

    // Show add member form
    'function showAddMemberForm(){' +
    '  document.getElementById("member-form-title").innerHTML="➕ Add New Member";' +
    '  document.getElementById("form-mode").value="add";' +
    '  document.getElementById("form-memberId").value="";' +
    '  document.getElementById("form-firstName").value="";' +
    '  document.getElementById("form-lastName").value="";' +
    '  document.getElementById("form-jobTitle").value="";' +
    '  document.getElementById("form-email").value="";' +
    '  document.getElementById("form-phone").value="";' +
    '  document.getElementById("form-location").value="";' +
    '  document.getElementById("form-unit").value="";' +
    '  document.getElementById("form-supervisor").value="";' +
    '  document.getElementById("form-manager").value="";' +
    '  document.getElementById("form-isSteward").value="No";' +
    '  var daysSelect=document.getElementById("form-officeDays");' +
    '  for(var i=0;i<daysSelect.options.length;i++)daysSelect.options[i].selected=false;' +
    '  document.getElementById("member-form-modal").style.display="block";' +
    '}' +

    // Ensure a select has the given value as an option, adding it if missing
    'function ensureOption(selectId,val){' +
    '  if(!val||val==="N/A")return;' +
    '  var sel=document.getElementById(selectId);if(!sel)return;' +
    '  for(var i=0;i<sel.options.length;i++){if(sel.options[i].value===val)return}' +
    '  sel.innerHTML+="<option value=\\""+escapeHtml(val)+"\\">"+escapeHtml(val)+"</option>";' +
    '}' +

    // Show edit member form with existing data
    'function showEditMemberForm(idx){' +
    '  var m=allMembers[idx];' +
    '  if(!m)return;' +
    '  document.getElementById("member-form-title").innerHTML="✏️ Edit Member: "+escapeHtml(m.name);' +
    '  document.getElementById("form-mode").value="edit";' +
    '  document.getElementById("form-memberId").value=m.id;' +
    '  document.getElementById("form-firstName").value=m.firstName||"";' +
    '  document.getElementById("form-lastName").value=m.lastName||"";' +
    '  var title=m.title!=="N/A"?m.title:"";' +
    '  ensureOption("form-jobTitle",title);' +
    '  document.getElementById("form-jobTitle").value=title;' +
    '  document.getElementById("form-email").value=m.email||"";' +
    '  document.getElementById("form-phone").value=m.phone||"";' +
    '  var loc=m.location!=="N/A"?m.location:"";' +
    '  ensureOption("form-location",loc);' +
    '  document.getElementById("form-location").value=loc;' +
    '  var unit=m.unit!=="N/A"?m.unit:"";' +
    '  ensureOption("form-unit",unit);' +
    '  document.getElementById("form-unit").value=unit;' +
    '  var sup=m.supervisor!=="N/A"?m.supervisor:"";' +
    '  ensureOption("form-supervisor",sup);' +
    '  document.getElementById("form-supervisor").value=sup;' +
    '  var mgr=m.manager!=="N/A"?m.manager:"";' +
    '  ensureOption("form-manager",mgr);' +
    '  document.getElementById("form-manager").value=mgr;' +
    '  document.getElementById("form-isSteward").value=m.isSteward?"Yes":"No";' +
    '  var daysSelect=document.getElementById("form-officeDays");' +
    '  var memberDays=m.officeDays&&m.officeDays!=="N/A"?m.officeDays.split(",").map(function(d){return d.trim()}):[];' +
    '  memberDays.forEach(function(d){if(d){var found=false;for(var j=0;j<daysSelect.options.length;j++){if(daysSelect.options[j].value===d){found=true;break}}if(!found)daysSelect.innerHTML+="<option value=\\""+escapeHtml(d)+"\\">"+escapeHtml(d)+"</option>"}});' +
    '  for(var i=0;i<daysSelect.options.length;i++){daysSelect.options[i].selected=memberDays.indexOf(daysSelect.options[i].value)>=0}' +
    '  document.getElementById("member-form-modal").style.display="block";' +
    '}' +

    // Close member form modal
    'function closeMemberForm(){' +
    '  document.getElementById("member-form-modal").style.display="none";' +
    '}' +

    // Save member (add or edit)
    'function saveMemberForm(){' +
    '  var mode=document.getElementById("form-mode").value;' +
    '  var firstName=document.getElementById("form-firstName").value.trim();' +
    '  var lastName=document.getElementById("form-lastName").value.trim();' +
    '  if(!firstName||!lastName){alert("First name and last name are required");return}' +
    '  var daysSelect=document.getElementById("form-officeDays");' +
    '  var selectedDays=[];' +
    '  for(var i=0;i<daysSelect.options.length;i++){if(daysSelect.options[i].selected)selectedDays.push(daysSelect.options[i].value)}' +
    '  var memberData={' +
    '    memberId:document.getElementById("form-memberId").value,' +
    '    firstName:firstName,' +
    '    lastName:lastName,' +
    '    jobTitle:document.getElementById("form-jobTitle").value,' +
    '    email:document.getElementById("form-email").value.trim(),' +
    '    phone:document.getElementById("form-phone").value.trim(),' +
    '    location:document.getElementById("form-location").value,' +
    '    unit:document.getElementById("form-unit").value,' +
    '    officeDays:selectedDays.join(", "),' +
    '    supervisor:document.getElementById("form-supervisor").value,' +
    '    manager:document.getElementById("form-manager").value,' +
    '    isSteward:document.getElementById("form-isSteward").value' +
    '  };' +
    '  var btn=document.querySelector("#member-form-modal .action-btn-primary");' +
    '  btn.disabled=true;btn.innerHTML="⏳ Saving...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result){' +
    '      btn.disabled=false;btn.innerHTML="💾 Save Member";' +
    '      closeMemberForm();' +
    '      alert(mode==="add"?"Member added successfully!":"Member updated successfully!");' +
    '      allMembers=[];loadMembers();' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      btn.disabled=false;btn.innerHTML="💾 Save Member";' +
    '      alert("Error saving member: "+e.message);' +
    '    })' +
    '    .saveInteractiveMember(memberData,mode);' +
    '}' +

    // Load grievances
    'function loadGrievances(){' +
    '  log("Loading grievances...");' +
    '  showLoading("grievances-list","Loading grievance log...");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){log("Grievances received:",data?data.length:0);allGrievances=data||[];renderGrievances(allGrievances)})'  +
    '    .withFailureHandler(function(e){logError("Failed to load grievances",e);document.getElementById("grievances-list").innerHTML="<div class=\\"error-state\\">Failed to load grievances<div class=\\"debug-info\\">"+escapeHtml(e.message)+"</div></div>"})' +
    '    .getInteractiveGrievanceData();' +
    '}' +

    // Render grievances with clickable details
    'function renderGrievances(data){' +
    '  var c=document.getElementById("grievances-list");' +
    '  if(!data||data.length===0){c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">📋</div><p>No grievances found</p></div>";return}' +
    '  c.innerHTML=data.slice(0,50).map(function(g,i){' +
    '    var badgeClass=g.isOverdue?"badge-overdue":(g.status==="Open"?"badge-open":(g.status==="Pending Info"?"badge-pending":"badge-closed"));' +
    '    var statusText=g.isOverdue?"Overdue":g.status;' +
    '    var daysInfo=g.isOverdue?"<span style=\\"color:#dc2626;font-weight:bold\\">⚠️ PAST DUE</span>":(typeof g.daysToDeadline==="number"?""+g.daysToDeadline+" days left":"");' +
    '    return "<div class=\\"list-item\\" onclick=\\"toggleGrievanceDetail(this,"+i+")\\">' +
    '      <div class=\\"list-item-header\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+escapeHtml(g.id)+" - "+escapeHtml(g.memberName)+"</div><div class=\\"list-item-subtitle\\">"+escapeHtml(g.issueType)+" • "+escapeHtml(g.currentStep)+"</div></div><div><span class=\\"badge "+badgeClass+"\\">"+escapeHtml(statusText)+"</span></div></div>' +
    '      <div class=\\"list-item-details\\">' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📅 Filed:</span><span class=\\"detail-value\\">"+escapeHtml(g.filedDate)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">🔔 Incident:</span><span class=\\"detail-value\\">"+escapeHtml(g.incidentDate)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">⏰ Next Due:</span><span class=\\"detail-value\\">"+escapeHtml(g.nextActionDue)+" "+daysInfo+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">⏱️ Days Open:</span><span class=\\"detail-value\\">"+escapeHtml(g.daysOpen)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📍 Location:</span><span class=\\"detail-value\\">"+escapeHtml(g.location)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📜 Articles:</span><span class=\\"detail-value\\">"+escapeHtml(g.articles)+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">🛡️ Steward:</span><span class=\\"detail-value\\">"+escapeHtml(g.steward)+"</span></div>' +
    '        "+(g.resolution?"<div class=\\"detail-row\\"><span class=\\"detail-label\\">✅ Resolution:</span><span class=\\"detail-value\\">"+escapeHtml(g.resolution)+"</span></div>":"")+"' +
    '        <div class=\\"detail-actions\\">' +
    '          <button class=\\"action-btn action-btn-primary\\" data-gid=\\""+escapeHtml(g.id)+"\\" onclick=\\"event.stopPropagation();google.script.run.showGrievanceQuickActions(this.dataset.gid)\\">⚡ Quick Actions</button>' +
    '          <button class=\\"action-btn action-btn-secondary\\" data-gid=\\""+escapeHtml(g.id)+"\\" onclick=\\"event.stopPropagation();google.script.run.navigateToGrievanceInSheet(this.dataset.gid)\\">📄 View in Sheet</button>' +
    '        </div>' +
    '      </div>' +
    '    </div>"' +
    '  }).join("");' +
    '  if(data.length>50)c.innerHTML+="<div class=\\"empty-state\\"><p>Showing 50 of "+data.length+" grievances. Use search/filters to find specific cases.</p></div>";' +
    '}' +

    // Toggle grievance detail
    'function toggleGrievanceDetail(el,idx){el.classList.toggle("expanded")}' +

    // Show specific grievance detail
    'function showGrievanceDetail(id){' +
    '  var g=allGrievances.find(function(x){return x.id===id});' +
    '  if(g){switchTab("grievances",document.getElementById("tab-grievances"));setTimeout(function(){' +
    '    document.getElementById("grievance-search").value=id;filterGrievances();' +
    '    var items=document.querySelectorAll("#grievances-list .list-item");if(items[0])items[0].classList.add("expanded");' +
    '  },300)}' +
    '}' +

    // Filter grievances
    'function filterGrievances(){' +
    '  var query=(document.getElementById("grievance-search").value||"").toLowerCase();' +
    '  var filtered=allGrievances;' +
    '  if(currentGrievanceFilter==="Overdue"){filtered=filtered.filter(function(g){return g.isOverdue})}' +
    '  else if(currentGrievanceFilter!=="all"){filtered=filtered.filter(function(g){return g.status===currentGrievanceFilter})}' +
    '  if(query&&query.length>=2){' +
    '    filtered=filtered.filter(function(g){' +
    '      return g.id.toLowerCase().indexOf(query)>=0||' +
    '             g.memberName.toLowerCase().indexOf(query)>=0||' +
    '             (g.issueType||"").toLowerCase().indexOf(query)>=0||' +
    '             (g.steward||"").toLowerCase().indexOf(query)>=0;' +
    '    });' +
    '  }' +
    '  renderGrievances(filtered);' +
    '}' +

    // Filter by status with button highlighting
    'function filterGrievanceStatus(status,btn){' +
    '  currentGrievanceFilter=status;' +
    '  document.querySelectorAll("#grievance-filter-bar .action-btn").forEach(function(b){' +
    '    b.classList.remove("active","action-btn-primary");' +
    '    if(b.dataset.filter!=="Overdue")b.classList.add("action-btn-secondary");' +
    '  });' +
    '  if(btn){btn.classList.add("active");if(status!=="Overdue")btn.classList.add("action-btn-primary");btn.classList.remove("action-btn-secondary")}' +
    '  filterGrievances();' +
    '}' +

    // Load analytics
    'function loadAnalytics(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){safeRun(function(){renderAnalytics(data)})})'  +
    '    .withFailureHandler(function(e){document.getElementById("analytics-charts").innerHTML="<div class=\\"error-state\\">Failed to load analytics</div>"})' +
    '    .getInteractiveAnalyticsData();' +
    '}' +

    // Load resources/links
    'function loadResources(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){resourceLinks=data||{};renderResources(data)})'  +
    '    .withFailureHandler(function(e){document.getElementById("resources-content").innerHTML="<div class=\\"error-state\\">Failed to load links</div>"})' +
    '    .getInteractiveResourceLinks();' +
    '}' +

    // Render analytics - ENHANCED VERSION
    'function renderAnalytics(data){' +
    '  var c=document.getElementById("analytics-charts");' +
    '  var html="";' +
    '  var colors=["#7C3AED","#059669","#1a73e8","#F97316","#DC2626","#8B5CF6","#10B981","#3B82F6","#F59E0B","#EF4444"];' +

    // ========== GRIEVANCE STATS SECTION ==========
    '  html+="<div class=\\"chart-container\\" style=\\"border-left:4px solid #7C3AED\\"><div class=\\"chart-title\\">📊 Grievance Statistics</div>";' +
    '  html+="<div class=\\"stats-grid\\" style=\\"margin-bottom:15px\\">";' +
    '  var totalG=data.statusCounts.open+data.statusCounts.pending+data.statusCounts.closed;' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+totalG+"</div><div class=\\"stat-label\\">Total Cases</div></div>";' +
    '  html+="<div class=\\"stat-card red\\"><div class=\\"stat-value\\">"+data.statusCounts.open+"</div><div class=\\"stat-label\\">Open</div></div>";' +
    '  html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+data.statusCounts.pending+"</div><div class=\\"stat-label\\">Pending</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+(data.grievanceStats?data.grievanceStats.avgDaysToResolve:0)+"</div><div class=\\"stat-label\\">Avg Days to Resolve</div></div>";' +
    '  html+="</div></div>";' +

    // ========== GRIEVANCES BY STATUS (with colors and proportional bars) ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📈 Grievances by Status</div><div class=\\"bar-chart\\">";' +
    '  if(data.grievanceStats&&data.grievanceStats.byStatus&&data.grievanceStats.byStatus.length>0){' +
    '    var maxStatus=Math.max.apply(null,data.grievanceStats.byStatus.map(function(s){return s.count}))||1;' +
    '    data.grievanceStats.byStatus.forEach(function(status){' +
    '      var pct=(status.count/maxStatus*100);' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:100px\\">"+escapeHtml(status.name)+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+status.color+"\\"></div></div><div class=\\"bar-value\\">"+status.count+"</div></div>";' +
    '    });' +
    '  }else if(totalG>0){' +
    '    var maxS=Math.max(data.statusCounts.open,data.statusCounts.pending,data.statusCounts.closed)||1;' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:100px\\">Open</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(data.statusCounts.open/maxS*100)+"%;background:#DC2626\\"></div></div><div class=\\"bar-value\\">"+data.statusCounts.open+"</div></div>";' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:100px\\">Pending</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(data.statusCounts.pending/maxS*100)+"%;background:#F97316\\"></div></div><div class=\\"bar-value\\">"+data.statusCounts.pending+"</div></div>";' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:100px\\">Closed</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(data.statusCounts.closed/maxS*100)+"%;background:#059669\\"></div></div><div class=\\"bar-value\\">"+data.statusCounts.closed+"</div></div>";' +
    '  }else{html+="<div class=\\"empty-state\\">No grievances</div>"}' +
    '  html+="</div></div>";' +

    // ========== GRIEVANCES BY TYPE (Issue Categories) ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📋 Grievances by Type (Issue Categories)</div><div class=\\"bar-chart\\">";' +
    '  var catData=data.grievanceStats&&data.grievanceStats.byType?data.grievanceStats.byType:data.topCategories;' +
    '  if(catData&&catData.length>0){' +
    '    var maxCat=Math.max.apply(null,catData.map(function(c){return c.count}))||1;' +
    '    catData.forEach(function(cat,idx){' +
    '      var pct=(cat.count/maxCat*100);' +
    '      var clr=colors[idx%colors.length];' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:130px;font-size:11px\\">"+escapeHtml(cat.name)+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+clr+"\\"></div></div><div class=\\"bar-value\\">"+cat.count+"</div></div>";' +
    '    });' +
    '  }else{html+="<div class=\\"empty-state\\">No issue data</div>"}' +
    '  html+="</div></div>";' +

    // ========== LOCATION BREAKDOWN ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📍 Grievances by Location</div><div class=\\"bar-chart\\">";' +
    '  if(data.grievanceStats&&data.grievanceStats.byLocation&&data.grievanceStats.byLocation.length>0){' +
    '    var maxLoc=Math.max.apply(null,data.grievanceStats.byLocation.map(function(l){return l.total}))||1;' +
    '    data.grievanceStats.byLocation.forEach(function(loc,idx){' +
    '      var pct=(loc.total/maxLoc*100);' +
    '      var clr=colors[idx%colors.length];' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:130px;font-size:11px\\">"+escapeHtml(loc.name)+"</div><div class=\\"bar-container\\" style=\\"position:relative\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+clr+"\\"></div>"+(loc.open>0?"<div style=\\"position:absolute;right:8px;top:2px;font-size:9px;color:#dc2626\\">"+loc.open+" open</div>":"")+"</div><div class=\\"bar-value\\">"+loc.total+"</div></div>";' +
    '    });' +
    '  }else{html+="<div class=\\"empty-state\\">No location data</div>"}' +
    '  html+="</div></div>";' +

    // ========== MONTH OVER MONTH TRENDS ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📅 Month Over Month Trends</div>";' +
    '  if(data.grievanceStats&&data.grievanceStats.monthlyTrends&&data.grievanceStats.monthlyTrends.length>0){' +
    '    html+="<div style=\\"display:flex;gap:5px;justify-content:space-around;margin:15px 0\\">";' +
    '    var maxMo=Math.max.apply(null,data.grievanceStats.monthlyTrends.map(function(m){return Math.max(m.filed,m.resolved)}))||1;' +
    '    data.grievanceStats.monthlyTrends.forEach(function(mo){' +
    '      var filedH=Math.max(mo.filed/maxMo*80,5);' +
    '      var resolvedH=Math.max(mo.resolved/maxMo*80,5);' +
    '      html+="<div style=\\"text-align:center;flex:1\\"><div style=\\"display:flex;gap:2px;justify-content:center;align-items:flex-end;height:90px\\">";' +
    '      html+="<div style=\\"width:16px;background:#DC2626;height:"+filedH+"px;border-radius:3px 3px 0 0\\" title=\\"Filed: "+mo.filed+"\\"></div>";' +
    '      html+="<div style=\\"width:16px;background:#059669;height:"+resolvedH+"px;border-radius:3px 3px 0 0\\" title=\\"Resolved: "+mo.resolved+"\\"></div>";' +
    '      html+="</div><div style=\\"font-size:10px;color:#666;margin-top:4px\\">"+mo.month.split("-")[1]+"/"+mo.month.split("-")[0].slice(2)+"</div></div>";' +
    '    });' +
    '    html+="</div><div style=\\"display:flex;justify-content:center;gap:15px;font-size:11px;color:#666\\"><span><span style=\\"display:inline-block;width:10px;height:10px;background:#DC2626;border-radius:2px\\"></span> Filed</span><span><span style=\\"display:inline-block;width:10px;height:10px;background:#059669;border-radius:2px\\"></span> Resolved</span></div>";' +
    '  }else{html+="<div class=\\"empty-state\\">No trend data available</div>"}' +
    '  html+="</div>";' +

    // ========== TOP 10 PERFORMERS BY SCORE ==========
    '  html+="<div class=\\"chart-container\\" style=\\"border-left:4px solid #059669\\"><div class=\\"chart-title\\">🏆 Top 10 Performers by Score</div>";' +
    '  if(data.stewardPerformance&&data.stewardPerformance.topPerformers&&data.stewardPerformance.topPerformers.length>0){' +
    '    html+="<table style=\\"width:100%;border-collapse:collapse;font-size:12px\\"><tr style=\\"background:#f3f4f6\\"><th style=\\"padding:8px;text-align:left\\">Rank</th><th style=\\"padding:8px;text-align:left\\">Steward</th><th style=\\"padding:8px;text-align:center\\">Score</th><th style=\\"padding:8px;text-align:center\\">Win Rate</th><th style=\\"padding:8px;text-align:center\\">Avg Days</th></tr>";' +
    '    data.stewardPerformance.topPerformers.forEach(function(p,i){' +
    '      var medal=i===0?"🥇":i===1?"🥈":i===2?"🥉":"";' +
    '      var scoreColor=p.score>=70?"#059669":p.score>=50?"#F97316":"#DC2626";' +
    '      html+="<tr style=\\"border-bottom:1px solid #e5e7eb\\"><td style=\\"padding:8px\\">"+medal+(i+1)+"</td><td style=\\"padding:8px\\">"+escapeHtml(p.name)+"</td><td style=\\"padding:8px;text-align:center;font-weight:bold;color:"+scoreColor+"\\">"+Math.round(p.score)+"</td><td style=\\"padding:8px;text-align:center\\">"+(p.winRate||0)+"%</td><td style=\\"padding:8px;text-align:center\\">"+(p.avgDays||0)+"</td></tr>";' +
    '    });' +
    '    html+="</table>";' +
    '  }else{html+="<div class=\\"empty-state\\">No performance data available.<br><small>Run Data Integrity Check to generate scores.</small></div>"}' +
    '  html+="</div>";' +

    // ========== TOP 10 BUSIEST STEWARDS ==========
    '  html+="<div class=\\"chart-container\\" style=\\"border-left:4px solid #F97316\\"><div class=\\"chart-title\\">📊 Top 10 Busiest Stewards (Case Load)</div>";' +
    '  if(data.stewardPerformance&&data.stewardPerformance.busiestStewards&&data.stewardPerformance.busiestStewards.length>0){' +
    '    var maxCases=Math.max.apply(null,data.stewardPerformance.busiestStewards.map(function(s){return s.total}))||1;' +
    '    html+="<div class=\\"bar-chart\\">";' +
    '    data.stewardPerformance.busiestStewards.forEach(function(s,idx){' +
    '      var pct=(s.total/maxCases*100);' +
    '      var openPct=(s.open/s.total*100);' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:110px;font-size:11px\\">"+escapeHtml(s.name)+"</div><div class=\\"bar-container\\" style=\\"position:relative\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:linear-gradient(90deg,#F97316 "+openPct+"%,#059669 "+openPct+"%)\\"></div></div><div class=\\"bar-value\\">"+s.total+" <small style=\\"color:#F97316\\">("+s.open+" open)</small></div></div>";' +
    '    });' +
    '    html+="</div>";' +
    '  }else{html+="<div class=\\"empty-state\\">No steward case data</div>"}' +
    '  html+="</div>";' +

    // ========== SURVEY RESULTS SECTION ==========
    '  html+="<div class=\\"chart-container\\" style=\\"border-left:4px solid #1a73e8\\"><div class=\\"chart-title\\">📊 Survey Results</div>";' +
    '  if(data.surveyResults){' +
    '    html+="<div class=\\"stats-grid\\" style=\\"margin-bottom:15px\\">";' +
    '    html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.surveyResults.totalResponses+"</div><div class=\\"stat-label\\">Total Responses</div></div>";' +
    '    html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+(data.surveyResults.avgSatisfaction||"-")+"</div><div class=\\"stat-label\\">Avg Satisfaction (1-10)</div></div>";' +
    '    html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+(data.surveyResults.responseRate||0)+"%</div><div class=\\"stat-label\\">Response Rate</div></div>";' +
    '    html+="</div>";' +
    '    if(data.surveyResults.bySection&&data.surveyResults.bySection.length>0){' +
    '      html+="<div class=\\"chart-title\\" style=\\"font-size:12px;margin:10px 0\\">Satisfaction by Section</div><div class=\\"bar-chart\\">";' +
    '      data.surveyResults.bySection.forEach(function(sec,idx){' +
    '        var pct=(sec.avg/10*100);' +
    '        var clr=sec.avg>=7?"#059669":sec.avg>=5?"#F97316":"#DC2626";' +
    '        html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:130px;font-size:11px\\">"+escapeHtml(sec.name)+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+clr+"\\"></div></div><div class=\\"bar-value\\">"+sec.avg+"</div></div>";' +
    '      });' +
    '      html+="</div>";' +
    '    }else{html+="<div style=\\"color:#999;font-size:12px;text-align:center;padding:10px\\">No section data. Complete surveys to see breakdown.</div>"}' +
    '  }else{html+="<div class=\\"empty-state\\">No survey data. Link Google Form to collect responses.</div>"}' +
    '  html+="</div>";' +

    // ========== RESOLUTION SUMMARY ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🏆 Resolution Summary</div>";' +
    '  html+="<div class=\\"stats-grid\\" style=\\"margin:0\\">";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.resolutions.won+"</div><div class=\\"stat-label\\">Won</div></div>";' +
    '  html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+data.resolutions.settled+"</div><div class=\\"stat-label\\">Settled</div></div>";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.resolutions.withdrawn+"</div><div class=\\"stat-label\\">Withdrawn</div></div>";' +
    '  html+="<div class=\\"stat-card red\\"><div class=\\"stat-value\\">"+data.resolutions.denied+"</div><div class=\\"stat-label\\">Denied</div></div>";' +
    '  html+="</div></div>";' +

    // ========== MEMBER DIRECTORY STATS ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">👥 Member Directory Statistics</div>";' +
    '  html+="<div class=\\"stats-grid\\" style=\\"margin-bottom:15px\\">";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.memberStats.total+"</div><div class=\\"stat-label\\">Total Members</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.memberStats.stewards+"</div><div class=\\"stat-label\\">Stewards</div></div>";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.memberStats.withOpenGrievance+"</div><div class=\\"stat-label\\">With Open Case</div></div>";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.memberStats.stewardRatio+"</div><div class=\\"stat-label\\">Member:Steward</div></div>";' +
    '  html+="</div></div>";' +

    // ========== MEMBERS BY LOCATION (improved visualization) ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📍 Members by Location</div>";' +
    '  if(data.memberStats.byLocation&&data.memberStats.byLocation.length>0){' +
    '    var maxLoc=Math.max.apply(null,data.memberStats.byLocation.map(function(l){return l.count}))||1;' +
    '    var totalMembers=data.memberStats.total||1;' +
    '    html+="<div style=\\"display:flex;flex-wrap:wrap;gap:8px;justify-content:center\\">";' +
    '    data.memberStats.byLocation.forEach(function(loc,idx){' +
    '      var pct=Math.round(loc.count/totalMembers*100);' +
    '      var size=Math.max(60,Math.min(120,60+(loc.count/maxLoc*60)));' +
    '      var clr=colors[idx%colors.length];' +
    '      html+="<div style=\\"text-align:center;padding:10px\\"><div style=\\"width:"+size+"px;height:"+size+"px;border-radius:50%;background:"+clr+";display:flex;flex-direction:column;align-items:center;justify-content:center;color:white;font-weight:bold;margin:0 auto\\"><span style=\\"font-size:"+(size/3)+"px\\">"+loc.count+"</span><span style=\\"font-size:10px\\">"+pct+"%</span></div><div style=\\"font-size:11px;color:#666;margin-top:6px;max-width:100px;overflow:hidden;text-overflow:ellipsis\\">"+escapeHtml(loc.name)+"</div></div>";' +
    '    });' +
    '    html+="</div>";' +
    '  }else{html+="<div class=\\"empty-state\\">No location data</div>"}' +
    '  html+="</div>";' +

    // ========== SANKEY DIAGRAM ==========
    '  var totalGrievances=data.statusCounts.open+data.statusCounts.pending+data.statusCounts.closed;' +
    '  if(totalGrievances>0){' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🔀 Grievance Flow</div>";' +
    '  html+="<div class=\\"sankey-container\\">";' +
    '  html+="<div class=\\"sankey-nodes\\">";' +
    '  html+="<div class=\\"sankey-column\\"><div class=\\"sankey-node source\\">Filed<br/>"+totalGrievances+"</div><div class=\\"sankey-label\\">Total Filed</div></div>";' +
    '  html+="<div class=\\"sankey-column\\">";' +
    '  if(data.statusCounts.open>0)html+="<div class=\\"sankey-node status-open\\">Open<br/>"+data.statusCounts.open+"</div>";' +
    '  if(data.statusCounts.pending>0)html+="<div class=\\"sankey-node status-pending\\">Pending<br/>"+data.statusCounts.pending+"</div>";' +
    '  if(data.statusCounts.closed>0)html+="<div class=\\"sankey-node status-closed\\">Closed<br/>"+data.statusCounts.closed+"</div>";' +
    '  html+="<div class=\\"sankey-label\\">Current Status</div></div>";' +
    '  html+="<div class=\\"sankey-column\\">";' +
    '  var totalResolved=data.resolutions.won+data.resolutions.settled+data.resolutions.withdrawn+data.resolutions.denied;' +
    '  if(data.resolutions.won>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#059669,#10b981)\\">Won "+data.resolutions.won+"</div>";' +
    '  if(data.resolutions.settled>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#f97316,#fb923c)\\">Settled "+data.resolutions.settled+"</div>";' +
    '  if(data.resolutions.withdrawn>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#6b7280,#9ca3af)\\">Withdrawn "+data.resolutions.withdrawn+"</div>";' +
    '  if(data.resolutions.denied>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#dc2626,#ef4444)\\">Denied "+data.resolutions.denied+"</div>";' +
    '  if(totalResolved===0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:#ccc\\">Pending</div>";' +
    '  html+="<div class=\\"sankey-label\\">Outcome</div></div>";' +
    '  html+="</div></div></div>";' +
    '  }' +
    '  c.innerHTML=html;' +
    '}' +

    // Render resources/links tab
    'function renderResources(data){' +
    '  var c=document.getElementById("resources-content");' +
    '  var html="";' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📝 Forms & Submissions</div><div class=\\"link-grid\\">";' +
    '  if(data.grievanceForm)html+="<a href=\\""+safeUrl(data.grievanceForm)+"\\" target=\\"_blank\\" class=\\"resource-link\\">📋 Grievance Form</a>";' +
    '  if(data.contactForm)html+="<a href=\\""+safeUrl(data.contactForm)+"\\" target=\\"_blank\\" class=\\"resource-link\\">✉️ Contact Form</a>";' +
    '  if(data.satisfactionForm)html+="<a href=\\""+safeUrl(data.satisfactionForm)+"\\" target=\\"_blank\\" class=\\"resource-link\\">📊 Satisfaction Survey</a>";' +
    '  if(!data.grievanceForm&&!data.contactForm&&!data.satisfactionForm)html+="<div class=\\"empty-state\\">No forms configured. Add URLs in Config sheet.</div>";' +
    '  html+="</div></div>";' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📂 Data & Documents</div><div class=\\"link-grid\\">";' +
    '  html+="<a href=\\""+safeUrl(data.spreadsheetUrl)+"\\" target=\\"_blank\\" class=\\"resource-link\\">📊 Open Full Spreadsheet</a>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showMemberDirectory()\\">👥 Member Directory</button>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showGrievanceLog()\\">📋 Grievance Log</button>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showConfigSheet()\\">⚙️ Configuration</button>";' +
    '  html+="</div></div>";' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🌐 External Links</div><div class=\\"link-grid\\">";' +
    '  if(data.orgWebsite)html+="<a href=\\""+safeUrl(data.orgWebsite)+"\\" target=\\"_blank\\" class=\\"resource-link\\">🏛️ Organization Website</a>";' +
    '  if(data.githubRepo)html+="<a href=\\""+safeUrl(data.githubRepo)+"\\" target=\\"_blank\\" class=\\"resource-link\\">📦 GitHub Repository</a>";' +
    '  html+="</div></div>";' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">⚡ Quick Actions</div><div class=\\"link-grid\\">";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showMobileUnifiedSearch()\\">🔍 Search All</button>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showMobileGrievanceForm()\\">➕ New Grievance</button>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showMyAssignedGrievances()\\">👤 My Cases</button>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showMemberSatisfactionDashboard()\\">📈 Satisfaction Dashboard</button>";' +
    '  html+="</div></div>";' +
    '  c.innerHTML=html;' +
    '}' +

    // Initialize
    'loadOverview();' +
    '</script>' +

    '</body></html>';
}

/**
 * Get overview data for interactive dashboard - ENHANCED with location data
 */
function getInteractiveOverviewData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = {
    totalMembers: 0,
    activeStewards: 0,
    totalGrievances: 0,
    openGrievances: 0,
    pendingInfo: 0,
    winRate: '0%',
    byLocation: []  // NEW: Location breakdown for overview
  };

  var locationMap = {};

  // Get member stats - only count rows with valid member IDs (starting with M)
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, MEMBER_COLS.IS_STEWARD).getValues();
    memberData.forEach(function(row) {
      var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
      // Skip blank rows - must have a valid member ID starting with M
      if (!memberId || (typeof memberId === 'string' && !memberId.toString().match(/^M/i))) return;

      data.totalMembers++;
      if (isTruthyValue(row[MEMBER_COLS.IS_STEWARD - 1])) data.activeStewards++;

      // Count by location
      var location = row[MEMBER_COLS.WORK_LOCATION - 1] || 'Unknown';
      if (!locationMap[location]) locationMap[location] = 0;
      locationMap[location]++;
    });

    // Convert location map to sorted array (top 8)
    data.byLocation = Object.keys(locationMap).map(function(key) {
      return { name: key, count: locationMap[key] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 8);
  }

  // Get grievance stats - only count rows with valid grievance IDs (starting with G)
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var grievanceData = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, GRIEVANCE_COLS.RESOLUTION).getValues();
    var wonCount = 0;
    var closedCount = 0;
    grievanceData.forEach(function(row) {
      var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
      // Skip blank rows - must have a valid grievance ID starting with G
      if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return;

      data.totalGrievances++;
      var status = row[GRIEVANCE_COLS.STATUS - 1];
      var resolution = row[GRIEVANCE_COLS.RESOLUTION - 1] || '';
      if (status === 'Open') data.openGrievances++;
      if (status === 'Pending Info') data.pendingInfo++;
      if (status !== 'Open' && status !== 'Pending Info') closedCount++;
      if (resolution.toLowerCase().indexOf('won') >= 0 || resolution.toLowerCase().indexOf('favorable') >= 0) wonCount++;
    });
    if (closedCount > 0) {
      data.winRate = Math.round(wonCount / closedCount * 100) + '%';
    }
  }

  return data;
}

/**
 * Get dropdown values from Config sheet for dashboard form population
 * @returns {Object} Map of dropdown field names to their available values
 */
function getConfigDropdownValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) return {};

  var result = {};
  var fields = [
    { key: 'jobTitles', col: CONFIG_COLS.JOB_TITLES },
    { key: 'locations', col: CONFIG_COLS.OFFICE_LOCATIONS },
    { key: 'units', col: CONFIG_COLS.UNITS },
    { key: 'officeDays', col: CONFIG_COLS.OFFICE_DAYS },
    { key: 'supervisors', col: CONFIG_COLS.SUPERVISORS },
    { key: 'managers', col: CONFIG_COLS.MANAGERS },
    { key: 'stewards', col: CONFIG_COLS.STEWARDS }
  ];

  var lastRow = configSheet.getLastRow();
  if (lastRow < 3) return result;

  var numRows = lastRow - 2;
  fields.forEach(function(f) {
    var data = configSheet.getRange(3, f.col, numRows, 1).getValues();
    result[f.key] = [];
    for (var i = 0; i < data.length; i++) {
      var val = data[i][0];
      if (val && val.toString().trim() !== '') {
        result[f.key].push(val.toString().trim());
      }
    }
  });

  return result;
}

/**
 * Get member data for interactive dashboard (expanded with more details)
 */
function getInteractiveMemberData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  ensureMinimumColumns(sheet, getMemberHeaders().length);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEMBER_COLS.QUICK_ACTIONS).getValues();
  return data.map(function(row) {
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
      officeDays: row[MEMBER_COLS.OFFICE_DAYS - 1] || 'N/A',
      email: row[MEMBER_COLS.EMAIL - 1] || '',
      phone: row[MEMBER_COLS.PHONE - 1] || '',
      preferredComm: row[MEMBER_COLS.PREFERRED_COMM - 1] || 'N/A',
      supervisor: row[MEMBER_COLS.SUPERVISOR - 1] || 'N/A',
      manager: row[MEMBER_COLS.MANAGER - 1] || 'N/A',
      isSteward: isTruthyValue(row[MEMBER_COLS.IS_STEWARD - 1]),
      assignedSteward: row[MEMBER_COLS.ASSIGNED_STEWARD - 1] || 'N/A',
      hasOpenGrievance: isTruthyValue(row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1]),
      grievanceStatus: row[MEMBER_COLS.GRIEVANCE_STATUS - 1] || ''
    };
  }).filter(function(m) { return m !== null; });
}

/**
 * Get grievance data for interactive dashboard (expanded with more details)
 */
function getInteractiveGrievanceData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  ensureMinimumColumns(sheet, getGrievanceHeaders().length);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.QUICK_ACTIONS).getValues();
  var tz = Session.getScriptTimeZone();

  return data.map(function(row, idx) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
    // Skip blank rows - must have a valid grievance ID starting with G
    if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return null;

    var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
    var incident = row[GRIEVANCE_COLS.INCIDENT_DATE - 1];
    var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

    return {
      id: grievanceId,
      rowNum: idx + 2, // For navigation back to sheet
      memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1] || '',
      memberName: ((row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '')).trim(),
      status: row[GRIEVANCE_COLS.STATUS - 1] || 'Filed',
      currentStep: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Step I',
      issueType: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'N/A',
      articles: row[GRIEVANCE_COLS.ARTICLES - 1] || 'N/A',
      filedDate: filed instanceof Date ? Utilities.formatDate(filed, tz, 'MM/dd/yyyy') : (filed || 'N/A'),
      incidentDate: incident instanceof Date ? Utilities.formatDate(incident, tz, 'MM/dd/yyyy') : (incident || 'N/A'),
      nextActionDue: nextDue instanceof Date ? Utilities.formatDate(nextDue, tz, 'MM/dd/yyyy') : (nextDue || 'N/A'),
      daysToDeadline: daysToDeadline,
      isOverdue: daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0),
      daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || 0,
      location: row[GRIEVANCE_COLS.LOCATION - 1] || 'N/A',
      steward: row[GRIEVANCE_COLS.STEWARD - 1] || 'N/A',
      resolution: row[GRIEVANCE_COLS.RESOLUTION - 1] || ''
    };
  }).filter(function(g) { return g !== null; });
}

/**
 * Get steward's assigned grievances for My Cases tab
 * Returns grievances where current user is the assigned steward
 */
function getMyStewardCases() {
  var email = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  ensureMinimumColumns(sheet, getGrievanceHeaders().length);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.QUICK_ACTIONS).getValues();
  var tz = Session.getScriptTimeZone();

  // Also check Member Directory to get steward name for matching
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var userStewardName = '';
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, MEMBER_COLS.IS_STEWARD).getValues();
    for (var i = 0; i < memberData.length; i++) {
      var memberEmail = memberData[i][MEMBER_COLS.EMAIL - 1] || '';
      if (memberEmail.toLowerCase() === email.toLowerCase() && isTruthyValue(memberData[i][MEMBER_COLS.IS_STEWARD - 1])) {
        userStewardName = ((memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (memberData[i][MEMBER_COLS.LAST_NAME - 1] || '')).trim();
        break;
      }
    }
  }

  return data.map(function(row, idx) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
    // Skip blank rows
    if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return null;

    // Check if current user is the steward for this grievance
    var steward = row[GRIEVANCE_COLS.STEWARD - 1] || '';
    var isMyCase = false;

    // Match by exact email (case-insensitive)
    if (steward && steward.toLowerCase() === email.toLowerCase()) {
      isMyCase = true;
    }
    // Match by exact full name if we found the user's steward name (case-insensitive)
    if (!isMyCase && userStewardName && steward && steward.toLowerCase() === userStewardName.toLowerCase()) {
      isMyCase = true;
    }

    if (!isMyCase) return null;

    var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
    var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

    return {
      id: grievanceId,
      rowNum: idx + 2,
      memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1] || '',
      memberName: ((row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '')).trim(),
      status: row[GRIEVANCE_COLS.STATUS - 1] || 'Filed',
      currentStep: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Step I',
      issueType: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'N/A',
      articles: row[GRIEVANCE_COLS.ARTICLES - 1] || 'N/A',
      filedDate: filed instanceof Date ? Utilities.formatDate(filed, tz, 'MM/dd/yyyy') : (filed || 'N/A'),
      nextActionDue: nextDue instanceof Date ? Utilities.formatDate(nextDue, tz, 'MM/dd/yyyy') : (nextDue || 'N/A'),
      daysToDeadline: daysToDeadline,
      isOverdue: daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0),
      daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || 0,
      location: row[GRIEVANCE_COLS.LOCATION - 1] || 'N/A'
    };
  }).filter(function(g) { return g !== null; });
}

/**
 * Get analytics data for interactive dashboard - ENHANCED VERSION
 * Now includes: grievance stats, steward performance, survey results, trends
 */
function getInteractiveAnalyticsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var data = {
    memberStats: {
      total: 0,
      stewards: 0,
      withOpenGrievance: 0,
      stewardRatio: '0:0',
      byLocation: [],
      byUnit: []
    },
    statusCounts: { open: 0, pending: 0, closed: 0 },
    topCategories: [],
    resolutions: { won: 0, settled: 0, withdrawn: 0, denied: 0 },
    // NEW: Enhanced grievance stats
    grievanceStats: {
      avgDaysToResolve: 0,
      totalResolved: 0,
      byType: [],
      byStatus: [],
      byLocation: [],
      monthlyTrends: []
    },
    // NEW: Top performers and busiest stewards
    stewardPerformance: {
      topPerformers: [],
      busiestStewards: []
    },
    // NEW: Survey results
    surveyResults: {
      totalResponses: 0,
      avgSatisfaction: 0,
      responseRate: 0,
      bySection: []
    }
  };

  // Get Member Directory statistics
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var locationMap = {};
  var unitMap = {};

  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, MEMBER_COLS.HAS_OPEN_GRIEVANCE).getValues();

    memberData.forEach(function(row) {
      var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
      if (!memberId || (typeof memberId === 'string' && !memberId.toString().match(/^M/i))) return;

      data.memberStats.total++;
      if (isTruthyValue(row[MEMBER_COLS.IS_STEWARD - 1])) data.memberStats.stewards++;
      if (isTruthyValue(row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1])) data.memberStats.withOpenGrievance++;

      var location = row[MEMBER_COLS.WORK_LOCATION - 1] || 'Unknown';
      if (!locationMap[location]) locationMap[location] = { members: 0, grievances: 0, open: 0 };
      locationMap[location].members++;

      var unit = row[MEMBER_COLS.UNIT - 1] || 'Unknown';
      if (!unitMap[unit]) unitMap[unit] = 0;
      unitMap[unit]++;
    });

    if (data.memberStats.stewards > 0) {
      var ratio = Math.round(data.memberStats.total / data.memberStats.stewards);
      data.memberStats.stewardRatio = ratio + ':1';
    } else {
      data.memberStats.stewardRatio = 'N/A';
    }

    data.memberStats.byLocation = Object.keys(locationMap).map(function(key) {
      return { name: key, count: locationMap[key].members };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 10);

    data.memberStats.byUnit = Object.keys(unitMap).map(function(key) {
      return { name: key, count: unitMap[key] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);
  }

  // Get Grievance Log statistics - ENHANCED
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var rows = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, GRIEVANCE_COLS.STEWARD).getValues();
    var categoryMap = {};
    var statusMap = {};
    var stewardCaseCount = {};
    var grievanceLocationMap = {};
    var monthlyMap = {};
    var totalDaysToResolve = 0;
    var resolvedCount = 0;

    rows.forEach(function(row) {
      var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
      if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return;

      var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
      var category = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'Other';
      var resolution = (row[GRIEVANCE_COLS.RESOLUTION - 1] || '').toLowerCase();
      var steward = row[GRIEVANCE_COLS.STEWARD - 1] || 'Unassigned';
      var location = row[GRIEVANCE_COLS.LOCATION - 1] || 'Unknown';
      var dateFiled = row[GRIEVANCE_COLS.DATE_FILED - 1];
      var dateClosed = row[GRIEVANCE_COLS.DATE_CLOSED - 1];
      var daysOpen = row[GRIEVANCE_COLS.DAYS_OPEN - 1];

      // Status counts
      if (status === 'Open') {
        data.statusCounts.open++;
        if (!statusMap['Open']) statusMap['Open'] = { count: 0, color: '#DC2626' };
        statusMap['Open'].count++;
      } else if (status === 'Pending Info') {
        data.statusCounts.pending++;
        if (!statusMap['Pending Info']) statusMap['Pending Info'] = { count: 0, color: '#F97316' };
        statusMap['Pending Info'].count++;
      } else if (status === 'Resolved' || status === 'Closed' || status === 'Withdrawn') {
        data.statusCounts.closed++;
        if (!statusMap[status]) statusMap[status] = { count: 0, color: '#059669' };
        statusMap[status].count++;

        // Calculate days to resolve
        if (dateFiled instanceof Date && dateClosed instanceof Date) {
          var days = Math.round((dateClosed - dateFiled) / (1000 * 60 * 60 * 24));
          if (days >= 0) {
            totalDaysToResolve += days;
            resolvedCount++;
          }
        } else if (typeof daysOpen === 'number' && daysOpen > 0) {
          totalDaysToResolve += daysOpen;
          resolvedCount++;
        }
      } else if (status) {
        if (!statusMap[status]) statusMap[status] = { count: 0, color: '#6B7280' };
        statusMap[status].count++;
      }

      // Category counts
      if (!categoryMap[category]) categoryMap[category] = 0;
      categoryMap[category]++;

      // Steward case counts
      if (steward && steward !== 'Unassigned') {
        if (!stewardCaseCount[steward]) stewardCaseCount[steward] = { total: 0, open: 0 };
        stewardCaseCount[steward].total++;
        if (status === 'Open' || status === 'Pending Info') stewardCaseCount[steward].open++;
      }

      // Location grievance counts
      if (!grievanceLocationMap[location]) grievanceLocationMap[location] = { total: 0, open: 0 };
      grievanceLocationMap[location].total++;
      if (status === 'Open' || status === 'Pending Info') grievanceLocationMap[location].open++;

      // Monthly trends
      if (dateFiled instanceof Date) {
        var monthKey = Utilities.formatDate(dateFiled, Session.getScriptTimeZone(), 'yyyy-MM');
        if (!monthlyMap[monthKey]) monthlyMap[monthKey] = { filed: 0, resolved: 0 };
        monthlyMap[monthKey].filed++;
      }
      if (dateClosed instanceof Date) {
        var closeMonthKey = Utilities.formatDate(dateClosed, Session.getScriptTimeZone(), 'yyyy-MM');
        if (!monthlyMap[closeMonthKey]) monthlyMap[closeMonthKey] = { filed: 0, resolved: 0 };
        monthlyMap[closeMonthKey].resolved++;
      }

      // Resolution counts
      if (resolution.indexOf('won') >= 0 || resolution.indexOf('favorable') >= 0) data.resolutions.won++;
      else if (resolution.indexOf('settled') >= 0) data.resolutions.settled++;
      else if (resolution.indexOf('withdrawn') >= 0) data.resolutions.withdrawn++;
      else if (resolution.indexOf('denied') >= 0 || resolution.indexOf('lost') >= 0) data.resolutions.denied++;
    });

    // Average days to resolve
    data.grievanceStats.avgDaysToResolve = resolvedCount > 0 ? Math.round(totalDaysToResolve / resolvedCount) : 0;
    data.grievanceStats.totalResolved = resolvedCount;

    // Top categories (issue types) - ALL of them, properly formatted
    data.topCategories = Object.keys(categoryMap).map(function(key) {
      return { name: key, count: categoryMap[key] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 10);

    // By type (same as categories but for the new section)
    data.grievanceStats.byType = data.topCategories;

    // By status with colors
    data.grievanceStats.byStatus = Object.keys(statusMap).map(function(key) {
      return { name: key, count: statusMap[key].count, color: statusMap[key].color };
    }).sort(function(a, b) { return b.count - a.count; });

    // Location breakdown for grievances
    data.grievanceStats.byLocation = Object.keys(grievanceLocationMap).map(function(key) {
      return { name: key, total: grievanceLocationMap[key].total, open: grievanceLocationMap[key].open };
    }).sort(function(a, b) { return b.total - a.total; }).slice(0, 10);

    // Monthly trends (last 6 months)
    var sortedMonths = Object.keys(monthlyMap).sort().slice(-6);
    data.grievanceStats.monthlyTrends = sortedMonths.map(function(key) {
      return { month: key, filed: monthlyMap[key].filed, resolved: monthlyMap[key].resolved };
    });

    // Top 10 busiest stewards
    data.stewardPerformance.busiestStewards = Object.keys(stewardCaseCount).map(function(key) {
      return { name: key, total: stewardCaseCount[key].total, open: stewardCaseCount[key].open };
    }).sort(function(a, b) { return b.total - a.total; }).slice(0, 10);
  }

  // Get Steward Performance data from hidden sheet
  var perfSheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);
  if (perfSheet && perfSheet.getLastRow() > 1) {
    try {
      var perfData = perfSheet.getRange(2, 1, Math.min(perfSheet.getLastRow() - 1, 20), STEWARD_PERF_COLS.PERFORMANCE_SCORE).getValues();
      data.stewardPerformance.topPerformers = perfData
        .filter(function(row) { return row[STEWARD_PERF_COLS.STEWARD - 1] && row[STEWARD_PERF_COLS.PERFORMANCE_SCORE - 1]; })
        .map(function(row) {
          return {
            name: row[STEWARD_PERF_COLS.STEWARD - 1],
            totalCases: row[STEWARD_PERF_COLS.TOTAL_CASES - 1] || 0,
            active: row[STEWARD_PERF_COLS.ACTIVE - 1] || 0,
            closed: row[STEWARD_PERF_COLS.CLOSED - 1] || 0,
            won: row[STEWARD_PERF_COLS.WON - 1] || 0,
            winRate: row[STEWARD_PERF_COLS.WIN_RATE - 1] || 0,
            avgDays: row[STEWARD_PERF_COLS.AVG_DAYS - 1] || 0,
            score: row[STEWARD_PERF_COLS.PERFORMANCE_SCORE - 1] || 0
          };
        })
        .sort(function(a, b) { return b.score - a.score; })
        .slice(0, 10);
    } catch (e) {
      Logger.log('Error reading steward performance: ' + e.message);
    }
  }

  // Get Survey Results from Member Satisfaction sheet (v4.23.0: dynamic col map)
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
  if (satSheet && satSheet.getLastRow() > 1) {
    try {
      var satLastRow   = satSheet.getLastRow();
      var numResponses = satLastRow - 1;
      data.surveyResults.totalResponses = numResponses;

      if (data.memberStats.total > 0) {
        data.surveyResults.responseRate = Math.round((numResponses / data.memberStats.total) * 100);
      }

      if (numResponses > 0) {
        // Use getSatisfactionSummary() for section averages — already dynamic
        var summary = getSatisfactionSummary();
        if (summary && summary.sections) {
          var overallSection = summary.sections['OVERALL_SAT'];
          if (overallSection && overallSection.avg !== null) {
            data.surveyResults.avgSatisfaction = overallSection.avg.toFixed(1);
          }
          var sectionOrder = ['OVERALL_SAT','STEWARD_3A','CHAPTER','LEADERSHIP','CONTRACT','COMMUNICATION','MEMBER_VOICE','VALUE_ACTION'];
          sectionOrder.forEach(function(key) {
            var sec = summary.sections[key];
            if (sec) {
              data.surveyResults.bySection.push({ name: sec.name, avg: sec.avg !== null ? parseFloat(sec.avg) : 0 });
            }
          });
        }
      }
    } catch(e) {
      Logger.log('Error reading satisfaction data: ' + e.message);
    }
  }

  return data;
}

/**
 * Get resource links from Config sheet for the dashboard
 */
function getInteractiveResourceLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  var links = {
    grievanceForm: '',
    contactForm: '',
    satisfactionForm: '',
    spreadsheetUrl: ss.getUrl(),
    orgWebsite: ''
  };

  if (configSheet && configSheet.getLastRow() > 1) {
    try {
      // Get URLs from Config sheet row 3 (data row; rows 1-2 are headers)
      // satisfactionForm removed v4.22.7 — survey is native webapp
      var row = configSheet.getRange(3, 1, 1, CONFIG_COLS.ORG_WEBSITE).getValues()[0];
      links.grievanceForm = row[CONFIG_COLS.GRIEVANCE_FORM_URL - 1] || '';
      links.contactForm = row[CONFIG_COLS.CONTACT_FORM_URL - 1] || '';
      links.orgWebsite = row[CONFIG_COLS.ORG_WEBSITE - 1] || '';
    } catch (e) {
      Logger.log('Error getting resource links: ' + e.message);
    }
  }

  return links;
}

/**
 * Get unique filter options for members (locations, units, office days)
 */
function getInteractiveMemberFilters() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  var filters = {
    locations: [],
    units: [],
    officeDays: []
  };

  if (configSheet && configSheet.getLastRow() > 1) {
    try {
      var lastRow = configSheet.getLastRow();
      if (lastRow < 3) return filters;
      var data = configSheet.getRange(3, 1, lastRow - 2, CONFIG_COLS.OFFICE_DAYS).getValues();

      // Get unique values from config
      data.forEach(function(row) {
        var loc = row[CONFIG_COLS.OFFICE_LOCATIONS - 1];
        var unit = row[CONFIG_COLS.UNITS - 1];
        var days = row[CONFIG_COLS.OFFICE_DAYS - 1];

        if (loc && filters.locations.indexOf(loc) === -1) filters.locations.push(loc);
        if (unit && filters.units.indexOf(unit) === -1) filters.units.push(unit);
        if (days && filters.officeDays.indexOf(days) === -1) filters.officeDays.push(days);
      });
    } catch (e) {
      Logger.log('Error getting filter options: ' + e.message);
    }
  }

  return filters;
}

/**
 * Navigate to a specific member in the Member Directory sheet
 */

/**
 * Navigate to a specific grievance in the Grievance Log sheet
 */

/**
 * Show the Member Directory sheet
 */

/**
 * Show the Grievance Log sheet
 */

/**
 * Show the Config sheet
 */

/**
 * Save a member from the interactive dashboard (add or edit)
 * @param {Object} memberData - Member data from the form
 * @param {string} mode - 'add' or 'edit'
 * @returns {Object} Result with success status
 */
function saveInteractiveMember(memberData, mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) throw new Error('Member Directory sheet not found');

  // --- Input validation ---
  var MAX_FIELD_LENGTH = 200;
  var EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var PHONE_REGEX = /^[\d\s\-\+\(\)\.ext]{0,30}$/;

  // Required fields
  if (!memberData.firstName || !String(memberData.firstName).trim()) {
    throw new Error('First name is required');
  }
  if (!memberData.lastName || !String(memberData.lastName).trim()) {
    throw new Error('Last name is required');
  }

  // Field length checks on all string fields
  var fieldsToCheck = ['firstName', 'lastName', 'jobTitle', 'location', 'unit',
    'officeDays', 'email', 'phone', 'supervisor', 'manager'];
  for (var f = 0; f < fieldsToCheck.length; f++) {
    var fieldName = fieldsToCheck[f];
    if (memberData[fieldName] && String(memberData[fieldName]).length > MAX_FIELD_LENGTH) {
      throw new Error(fieldName + ' exceeds maximum length of ' + MAX_FIELD_LENGTH + ' characters');
    }
  }

  // Email format validation (if provided)
  if (memberData.email && String(memberData.email).trim() !== '') {
    if (!EMAIL_REGEX.test(String(memberData.email).trim())) {
      throw new Error('Invalid email format');
    }
  }

  // Phone format validation (if provided)
  if (memberData.phone && String(memberData.phone).trim() !== '') {
    if (!PHONE_REGEX.test(String(memberData.phone).trim())) {
      throw new Error('Invalid phone format. Use digits, spaces, dashes, parentheses, or dots (max 30 chars)');
    }
  }

  // isSteward must be Yes or No
  if (memberData.isSteward && memberData.isSteward !== 'Yes' && memberData.isSteward !== 'No') {
    throw new Error('isSteward must be "Yes" or "No"');
  }

  // Sanitize all string values with escapeForFormula to prevent formula injection
  var safeFirstName = escapeForFormula(String(memberData.firstName || '').trim());
  var safeLastName = escapeForFormula(String(memberData.lastName || '').trim());
  var safeJobTitle = escapeForFormula(String(memberData.jobTitle || '').trim());
  var safeLocation = escapeForFormula(String(memberData.location || '').trim());
  var safeUnit = escapeForFormula(String(memberData.unit || '').trim());
  var safeOfficeDays = escapeForFormula(String(memberData.officeDays || '').trim());
  var safeEmail = escapeForFormula(String(memberData.email || '').trim());
  var safePhone = escapeForFormula(String(memberData.phone || '').trim());
  var safeSupervisor = escapeForFormula(String(memberData.supervisor || '').trim());
  var safeManager = escapeForFormula(String(memberData.manager || '').trim());
  var safeIsSteward = (memberData.isSteward === 'Yes') ? 'Yes' : 'No';

  if (mode === 'add') {
    // Generate a new member ID
    var existingIds = {};
    var idData = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, Math.max(1, sheet.getLastRow() - 1), 1).getValues();
    idData.forEach(function(row) {
      if (row[0]) existingIds[row[0]] = true;
    });

    var newId = generateNameBasedId('M', safeFirstName, safeLastName, existingIds);

    // Create new row array
    var newRow = [];
    for (var i = 0; i < MEMBER_COLS.QUICK_ACTIONS; i++) newRow.push('');

    newRow[MEMBER_COLS.MEMBER_ID - 1] = newId;
    newRow[MEMBER_COLS.FIRST_NAME - 1] = safeFirstName;
    newRow[MEMBER_COLS.LAST_NAME - 1] = safeLastName;
    newRow[MEMBER_COLS.JOB_TITLE - 1] = safeJobTitle;
    newRow[MEMBER_COLS.WORK_LOCATION - 1] = safeLocation;
    newRow[MEMBER_COLS.UNIT - 1] = safeUnit;
    newRow[MEMBER_COLS.OFFICE_DAYS - 1] = safeOfficeDays;
    newRow[MEMBER_COLS.EMAIL - 1] = safeEmail;
    newRow[MEMBER_COLS.PHONE - 1] = safePhone;
    newRow[MEMBER_COLS.SUPERVISOR - 1] = safeSupervisor;
    newRow[MEMBER_COLS.MANAGER - 1] = safeManager;
    newRow[MEMBER_COLS.IS_STEWARD - 1] = safeIsSteward;

    // Append the new row
    sheet.appendRow(newRow);
    ss.toast('New member added: ' + safeFirstName + ' ' + safeLastName + ' (' + newId + ')', 'Member Added', 5);

    return { success: true, memberId: newId, mode: 'add' };

  } else if (mode === 'edit') {
    // Find the member row by ID
    var memberId = memberData.memberId;
    if (!memberId) throw new Error('Member ID is required for editing');

    var data = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, sheet.getLastRow() - 1, 1).getValues();
    var rowIndex = -1;
    for (var j = 0; j < data.length; j++) {
      if (data[j][0] === memberId) {
        rowIndex = j + 2; // Row 1 is header
        break;
      }
    }

    if (rowIndex === -1) throw new Error('Member not found: ' + memberId);

    // Update the member data
    sheet.getRange(rowIndex, MEMBER_COLS.FIRST_NAME).setValue(safeFirstName);
    sheet.getRange(rowIndex, MEMBER_COLS.LAST_NAME).setValue(safeLastName);
    sheet.getRange(rowIndex, MEMBER_COLS.JOB_TITLE).setValue(safeJobTitle);
    sheet.getRange(rowIndex, MEMBER_COLS.WORK_LOCATION).setValue(safeLocation);
    sheet.getRange(rowIndex, MEMBER_COLS.UNIT).setValue(safeUnit);
    sheet.getRange(rowIndex, MEMBER_COLS.OFFICE_DAYS).setValue(safeOfficeDays);
    sheet.getRange(rowIndex, MEMBER_COLS.EMAIL).setValue(safeEmail);
    sheet.getRange(rowIndex, MEMBER_COLS.PHONE).setValue(safePhone);
    sheet.getRange(rowIndex, MEMBER_COLS.SUPERVISOR).setValue(safeSupervisor);
    sheet.getRange(rowIndex, MEMBER_COLS.MANAGER).setValue(safeManager);
    sheet.getRange(rowIndex, MEMBER_COLS.IS_STEWARD).setValue(safeIsSteward);

    ss.toast('Member updated: ' + safeFirstName + ' ' + safeLastName, 'Member Updated', 5);

    return { success: true, memberId: memberId, mode: 'edit' };
  }

  throw new Error('Invalid mode: ' + mode);
}

// ╔═══════════════════════════════════════════════════════════════════════════╗
// ║                                                                           ║
// ║         ⚠️  END OF PROTECTED SECTION - INTERACTIVE DASHBOARD  ⚠️         ║
// ║                                                                           ║
// ╚═══════════════════════════════════════════════════════════════════════════╝
/**
 * ============================================================================
 * DASHBOARD - STRATEGIC COMMAND CENTER MASTER ENGINE (V 3.6.0)
 * ============================================================================
 * CORE FEATURES:
 * 1. Dual-Dashboard Architecture:
 *    - Executive View (Internal, shows PII, Case Management)
 *    - Member Analytics (Expansive, No PII, Strategic Reporting)
 * 2. Visual KPIs: Success Gauges, Sentiment Radars, Participation Funnels.
 * 3. Strategic Intelligence: Hot Zone Heatmaps & Rising Star Identification.
 * 4. Automation: Midnight Auto-Refresh & Critical Performance Email Alerts.
 * 5. Peak Performance: Array-based batch processing for speed and reliability.
 * 6. Auto-ID Generator & Duplicate Prevention
 * 7. Legal Stage-Gate Workflow (Intake to Arbitration)
 * 8. Automatic Chief Steward Escalation Alerts
 * 9. Digital Signature Block PDF Generation
 * 10. Automated Global UI Styling (Roboto Theme & Status Colors)
 * ============================================================================
 *
 * OFFICIAL BRANDING PATHS:
 * - All automated emails use: "[Strategic Command Center] Status Update"
 * - All PDF metadata lists "Strategic Command Center" as the Author
 * ============================================================================
 */

