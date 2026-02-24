/**
 * ConfigReader.gs
 * Reads the Config tab (HORIZONTAL layout: headers in Row 1, values in Row 2+).
 * Some columns are single-value (Organization Name), some are lists (Job Titles).
 * All lookups by header name — never by column index.
 */

var ConfigReader = (function () {
  
  var CACHE_KEY = 'ORG_CONFIG';
  var CACHE_TTL = 21600;
  var CONFIG_SHEET_NAME = 'Config';
  
  var SINGLE_KEYS = {
    orgName:              ['organization name'],
    localNumber:          ['local number'],
    mainOfficeAddress:    ['main office address'],
    mainPhone:            ['main phone'],
    mainFax:              ['main fax'],
    mainContactName:      ['main contact name'],
    mainContactEmail:     ['main contact email'],
    chiefStewardEmail:    ['chief steward email'],
    googleDriveFolderId:  ['google drive folder id'],
    googleCalendarId:     ['google calendar id'],
    grievanceFormUrl:     ['grievance form url'],
    contactFormUrl:       ['contact form url'],
    satisfactionSurveyUrl:['satisfaction survey url'],
    mobileDashboardUrl:   ['\u{1F4F1} mobile dashboard url', 'mobile dashboard url'],
    contractName:         ['contract name'],
    unionParent:          ['union parent'],
    stateRegion:          ['state/region'],
    orgWebsite:           ['organization website'],
    archiveFolderId:      ['archive folder id'],
    templateId:           ['template id'],
    pdfFolderId:          ['pdf folder id'],
    filingDeadlineDays:   ['filing deadline days'],
    stepIResponseDays:    ['step i response days'],
    stepIIAppealDays:     ['step ii appeal days'],
    stepIIResponseDays:   ['step ii response days'],
    alertDaysBeforeDeadline: ['alert days before deadline'],
  };
  
  var LIST_KEYS = {
    jobTitles:            ['job titles'],
    officeLocations:      ['office locations'],
    units:                ['units'],
    officeDays:           ['office days'],
    supervisors:          ['supervisors'],
    managers:             ['managers'],
    stewards:             ['stewards'],
    stewardCommittees:    ['steward committees'],
    grievanceStatuses:    ['grievance status'],
    grievanceSteps:       ['grievance step'],
    issueCategories:      ['issue category'],
    articlesViolated:     ['articles violated'],
    communicationMethods: ['communication methods'],
    grievanceCoordinators:['grievance coordinators'],
    adminEmails:          ['admin emails'],
    notificationRecipients:['notification recipients'],
    bestTimesToContact:   ['best times to contact'],
    homeTowns:            ['home towns'],
    officeAddresses:      ['office addresses'],
    unitCodes:            ['unit codes'],
    escalationStatuses:   ['escalation statuses'],
    escalationSteps:      ['escalation steps'],
  };
  
  // Dashboard-specific (may need adding to Config tab)
  var DASH_KEYS = {
    accentHue:           ['accent hue', 'dashboard accent hue'],
    logoInitials:        ['logo initials'],
    magicLinkExpiryDays: ['magic link expiry days', 'magic link expiry'],
    cookieDurationDays:  ['cookie duration days', 'cookie duration'],
    stewardLabel:        ['steward label'],
    memberLabel:         ['member label'],
  };
  
  function getConfig(forceRefresh) {
    if (!forceRefresh) {
      try {
        var cached = CacheService.getScriptCache().get(CACHE_KEY);
        if (cached) return JSON.parse(cached);
      } catch (e) {}
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
    if (!sheet) throw new Error('Config tab not found.');
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) throw new Error('Config tab empty.');
    
    var colMap = {};
    for (var i = 0; i < data[0].length; i++) {
      var h = String(data[0][i]).trim().toLowerCase();
      if (h) colMap[h] = i;
    }
    
    var config = {};
    
    // Single values
    for (var key in SINGLE_KEYS) {
      config[key] = _single(data, colMap, SINGLE_KEYS[key]);
    }
    
    // List values
    for (var key in LIST_KEYS) {
      config[key] = _list(data, colMap, LIST_KEYS[key]);
    }
    
    // Dashboard settings with defaults
    config.accentHue = _int(_single(data, colMap, DASH_KEYS.accentHue), 250);
    config.logoInitials = _single(data, colMap, DASH_KEYS.logoInitials) || _initials(config.orgName);
    config.magicLinkExpiryDays = _int(_single(data, colMap, DASH_KEYS.magicLinkExpiryDays), 7);
    config.cookieDurationDays = _int(_single(data, colMap, DASH_KEYS.cookieDurationDays), 30);
    config.stewardLabel = _single(data, colMap, DASH_KEYS.stewardLabel) || 'Steward';
    config.memberLabel = _single(data, colMap, DASH_KEYS.memberLabel) || 'Member';
    
    // Derived
    config.magicLinkExpiryMs = config.magicLinkExpiryDays * 86400000;
    config.cookieDurationMs = config.cookieDurationDays * 86400000;
    config.orgAbbrev = config.localNumber || _initials(config.orgName);
    config.filingDeadlineDays = _int(config.filingDeadlineDays, 15);
    config.stepIResponseDays = _int(config.stepIResponseDays, 10);
    config.stepIIAppealDays = _int(config.stepIIAppealDays, 10);
    config.stepIIResponseDays = _int(config.stepIIResponseDays, 10);
    config.alertDaysBeforeDeadline = _int(config.alertDaysBeforeDeadline, 3);
    
    try { CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(config), CACHE_TTL); } catch (e) {}
    return config;
  }
  
  function getConfigJSON() { return JSON.stringify(getSafeConfig()); }
  
  function getSafeConfig() {
    var c = getConfig();
    return {
      orgName: c.orgName, orgAbbrev: c.orgAbbrev, localNumber: c.localNumber || '',
      logoInitials: c.logoInitials, accentHue: c.accentHue,
      stewardLabel: c.stewardLabel, memberLabel: c.memberLabel,
      magicLinkExpiryDays: c.magicLinkExpiryDays, cookieDurationDays: c.cookieDurationDays,
      grievanceFormUrl: c.grievanceFormUrl || '', contactFormUrl: c.contactFormUrl || '',
      orgWebsite: c.orgWebsite || '', contractName: c.contractName || '',
      mainPhone: c.mainPhone || '', unionParent: c.unionParent || '',
    };
  }
  
  function refreshConfig() { return getConfig(true); }
  
  function _findCol(colMap, aliases) {
    for (var i = 0; i < aliases.length; i++) { if (colMap.hasOwnProperty(aliases[i].toLowerCase())) return colMap[aliases[i].toLowerCase()]; }
    return -1;
  }
  function _single(data, colMap, aliases) {
    var c = _findCol(colMap, aliases); if (c === -1) return '';
    for (var r = 1; r < data.length; r++) { var v = data[r][c]; if (v !== null && v !== undefined && String(v).trim()) return String(v).trim(); }
    return '';
  }
  function _list(data, colMap, aliases) {
    var c = _findCol(colMap, aliases); if (c === -1) return [];
    var out = [];
    for (var r = 1; r < data.length; r++) { var v = data[r][c]; if (v !== null && v !== undefined && String(v).trim()) out.push(String(v).trim()); }
    return out;
  }
  function _int(v, d) { var p = parseInt(v, 10); return isNaN(p) ? d : p; }
  function _initials(n) { if (!n) return 'ORG'; return String(n).trim().split(/\s+/).map(function(w){return w[0];}).join('').substring(0,3).toUpperCase(); }
  
  return { getConfig: getConfig, getConfigJSON: getConfigJSON, getSafeConfig: getSafeConfig, refreshConfig: refreshConfig };
})();
