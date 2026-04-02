/**
 * Tests for 33_NewFeatureServices.gs
 *
 * Covers all IIFE service modules: HandoffService, MentorshipService,
 * CommunicationLogService, KnowledgeBaseService, DigestService,
 * DocumentChecklistService, EscalationEngine, ReportService,
 * TwoFactorService, SMSService, RSVPService, and their global wrappers.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Stub dependencies
global.logAuditEvent = jest.fn();
global.setSheetVeryHidden_ = jest.fn();
global._refreshNavBadges = jest.fn();
global.escapeHtml = jest.fn(function(s) { return String(s); });
global.safeSendEmail_ = jest.fn(function() { return { success: true }; });
global.getConfigValue_ = jest.fn(function() { return ''; });
global.recordSecurityEvent = jest.fn();
global.maskEmail = jest.fn(function(e) { return e; });
global.SECURITY_SEVERITY = { LOW: 'LOW', MEDIUM: 'MEDIUM', HIGH: 'HIGH', CRITICAL: 'CRITICAL' };
global.loadMemberData_ = jest.fn(function() { return { leaders: [] }; });

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '33_NewFeatureServices.gs']);

// ============================================================================
// Module existence — all services
// ============================================================================

describe('NewFeatureServices module existence', () => {
  test('HandoffService is defined with expected API', () => {
    expect(HandoffService).toBeDefined();
    expect(typeof HandoffService.initSheet).toBe('function');
    expect(typeof HandoffService.getHandoffNotes).toBe('function');
    expect(typeof HandoffService.addHandoffNote).toBe('function');
    expect(typeof HandoffService.archiveHandoffNote).toBe('function');
  });

  test('MentorshipService is defined with expected API', () => {
    expect(MentorshipService).toBeDefined();
    expect(typeof MentorshipService.initSheet).toBe('function');
    expect(typeof MentorshipService.getPairings).toBe('function');
    expect(typeof MentorshipService.createPairing).toBe('function');
    expect(typeof MentorshipService.updatePairingNotes).toBe('function');
    expect(typeof MentorshipService.closePairing).toBe('function');
    expect(typeof MentorshipService.suggestPairings).toBe('function');
  });

  test('CommunicationLogService is defined with expected API', () => {
    expect(CommunicationLogService).toBeDefined();
    expect(typeof CommunicationLogService.initSheet).toBe('function');
    expect(typeof CommunicationLogService.logCommunication).toBe('function');
    expect(typeof CommunicationLogService.getCommunicationLog).toBe('function');
    expect(typeof CommunicationLogService.getStewardCommunicationSummary).toBe('function');
  });

  test('KnowledgeBaseService is defined with expected API', () => {
    expect(KnowledgeBaseService).toBeDefined();
    expect(typeof KnowledgeBaseService.initSheet).toBe('function');
    expect(typeof KnowledgeBaseService.searchArticles).toBe('function');
    expect(typeof KnowledgeBaseService.getArticle).toBe('function');
    expect(typeof KnowledgeBaseService.getRelatedArticles).toBe('function');
    expect(typeof KnowledgeBaseService.addArticle).toBe('function');
  });

  test('DigestService is defined with expected API', () => {
    expect(DigestService).toBeDefined();
    expect(typeof DigestService.initSheet).toBe('function');
    expect(typeof DigestService.getPreferences).toBe('function');
    expect(typeof DigestService.setPreferences).toBe('function');
    expect(typeof DigestService.buildDigestContent).toBe('function');
    expect(typeof DigestService.sendScheduledDigests).toBe('function');
  });

  test('DocumentChecklistService is defined with expected API', () => {
    expect(DocumentChecklistService).toBeDefined();
    expect(typeof DocumentChecklistService.getChecklist).toBe('function');
    expect(typeof DocumentChecklistService.getCompletionStatus).toBe('function');
  });

  test('EscalationEngine is defined with expected API', () => {
    expect(EscalationEngine).toBeDefined();
    expect(typeof EscalationEngine.getRecommendation).toBe('function');
  });

  test('ReportService is defined with expected API', () => {
    expect(ReportService).toBeDefined();
    expect(typeof ReportService.generateMonthlyReport).toBe('function');
    expect(typeof ReportService.generateReportHtml).toBe('function');
  });

  test('TwoFactorService is defined with expected API', () => {
    expect(TwoFactorService).toBeDefined();
    expect(typeof TwoFactorService.generateCode).toBe('function');
    expect(typeof TwoFactorService.verifyCode).toBe('function');
    expect(typeof TwoFactorService.hasValidSession).toBe('function');
  });

  test('SMSService is defined with expected API', () => {
    expect(SMSService).toBeDefined();
    expect(typeof SMSService.initSheet).toBe('function');
    expect(typeof SMSService.configureProvider).toBe('function');
    expect(typeof SMSService.isConfigured).toBe('function');
    expect(typeof SMSService.getProviderStatus).toBe('function');
    expect(typeof SMSService.sendSMS).toBe('function');
    expect(typeof SMSService.sendStatusUpdate).toBe('function');
    expect(typeof SMSService.testSMS).toBe('function');
    expect(typeof SMSService.getLog).toBe('function');
  });

  test('RSVPService is defined with expected API', () => {
    expect(RSVPService).toBeDefined();
    expect(typeof RSVPService.initSheet).toBe('function');
    expect(typeof RSVPService.sendInvitations).toBe('function');
    expect(typeof RSVPService.processRSVP).toBe('function');
    expect(typeof RSVPService.getRSVPSummary).toBe('function');
    expect(typeof RSVPService.reconcileAttendance).toBe('function');
  });
});

// ============================================================================
// Global wrappers existence
// ============================================================================

describe('NewFeatureServices global wrappers', () => {
  var wrappers = [
    'dataGetHandoffNotes', 'dataAddHandoffNote', 'dataArchiveHandoffNote', 'dataInitHandoffNotes',
    'dataGetMentorshipPairings', 'dataCreateMentorshipPairing', 'dataUpdateMentorshipNotes',
    'dataCloseMentorshipPairing', 'dataGetMentorshipSuggestions', 'dataGetMemberLeaders',
    'dataLogCommunication', 'dataGetCommunicationLog',
    'dataSearchKnowledgeBase', 'dataGetKnowledgeBaseArticle', 'dataGetRelatedArticles', 'dataAddKnowledgeBaseArticle',
    'dataGetDigestPreferences', 'dataSetDigestPreferences', 'dataSendScheduledDigests',
    'dataGetDocumentChecklist',
    'dataGetEscalationRecommendation',
    'dataGenerateMonthlyReport', 'dataGenerateReportHtml',
    'dataRequest2FACode', 'dataVerify2FACode',
    'dataConfigureTwilio', 'dataSendSMS', 'dataGetSMSLog', 'dataTestSMS',
  ];

  wrappers.forEach(function(name) {
    test(name + ' is a function', () => {
      expect(typeof global[name]).toBe('function');
    });
  });
});

// ============================================================================
// HandoffService
// ============================================================================

describe('HandoffService.getHandoffNotes', () => {
  test('returns empty array when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    expect(HandoffService.getHandoffNotes('CASE-1')).toEqual([]);
  });

  test('returns empty array when sheet does not exist', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    expect(HandoffService.getHandoffNotes('CASE-1')).toEqual([]);
  });

  test('filters notes by caseId and excludes archived', () => {
    var HEADERS = ['ID', 'Case ID', 'From Steward', 'To Steward', 'Note Text', 'Created', 'Status'];
    var data = [
      HEADERS,
      ['N1', 'CASE-1', 'steward@test.com', 'other@test.com', 'Note one', new Date(), 'active'],
      ['N2', 'CASE-2', 'steward@test.com', '', 'Wrong case', new Date(), 'active'],
      ['N3', 'CASE-1', 'steward@test.com', '', 'Archived note', new Date(), 'archived'],
    ];
    var sheet = createMockSheet(SHEETS.HANDOFF_NOTES, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = HandoffService.getHandoffNotes('CASE-1');
    expect(result.length).toBe(1);
    expect(result[0].id).toBe('N1');
    expect(result[0].noteText).toBe('Note one');
  });
});

describe('HandoffService.addHandoffNote', () => {
  test('appends a row and returns success with id', () => {
    var sheet = createMockSheet(SHEETS.HANDOFF_NOTES, [['ID']]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = HandoffService.addHandoffNote('steward@test.com', 'CASE-1', 'My note', 'to@test.com');
    expect(result.success).toBe(true);
    expect(result.id).toBeTruthy();
    expect(sheet.appendRow).toHaveBeenCalled();
  });
});

// ============================================================================
// MentorshipService
// ============================================================================

describe('MentorshipService.createPairing', () => {
  test('rejects missing emails', () => {
    var result = MentorshipService.createPairing('', 'mentee@test.com');
    expect(result.success).toBe(false);
  });

  test('rejects invalid email format', () => {
    var result = MentorshipService.createPairing('not-an-email', 'mentee@test.com');
    expect(result.success).toBe(false);
    expect(result.message).toContain('Invalid email');
  });

  test('rejects same mentor and mentee', () => {
    var result = MentorshipService.createPairing('same@test.com', 'same@test.com');
    expect(result.success).toBe(false);
    expect(result.message).toContain('same person');
  });

  test('creates pairing with valid inputs', () => {
    var sheet = createMockSheet(SHEETS.MENTORSHIP, [['ID']]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = MentorshipService.createPairing('mentor@test.com', 'mentee@test.com', 'Discipline');
    expect(result.success).toBe(true);
    expect(result.id).toBeTruthy();
  });
});

describe('MentorshipService.getPairings', () => {
  test('returns empty array when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    expect(MentorshipService.getPairings()).toEqual([]);
  });

  test('excludes closed pairings', () => {
    var HEADERS = ['ID', 'Mentor Email', 'Mentee Email', 'Case Types', 'Status', 'Started', 'Notes'];
    var data = [
      HEADERS,
      ['P1', 'mentor@test.com', 'mentee1@test.com', 'Discipline', 'active', new Date(), ''],
      ['P2', 'mentor@test.com', 'mentee2@test.com', '', 'closed', new Date(), ''],
    ];
    var sheet = createMockSheet(SHEETS.MENTORSHIP, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = MentorshipService.getPairings();
    expect(result.length).toBe(1);
    expect(result[0].id).toBe('P1');
  });
});

// ============================================================================
// CommunicationLogService
// ============================================================================

describe('CommunicationLogService.getCommunicationLog', () => {
  test('returns empty array when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    expect(CommunicationLogService.getCommunicationLog('user@test.com')).toEqual([]);
  });

  test('returns entries for specified member email', () => {
    var HEADERS = ['ID', 'Member Email', 'Steward Email', 'Type', 'Subject', 'Notes', 'Timestamp'];
    var data = [
      HEADERS,
      ['C1', 'user@test.com', 'steward@test.com', 'phone', 'Check-in', 'Called about meeting', new Date()],
      ['C2', 'other@test.com', 'steward@test.com', 'email', 'FYI', 'Sent update', new Date()],
    ];
    var sheet = createMockSheet(SHEETS.COMMUNICATION_LOG, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = CommunicationLogService.getCommunicationLog('user@test.com');
    expect(result.length).toBe(1);
    expect(result[0].id).toBe('C1');
    expect(result[0].type).toBe('phone');
  });
});

describe('CommunicationLogService.getStewardCommunicationSummary', () => {
  test('returns zero counts when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    var result = CommunicationLogService.getStewardCommunicationSummary('steward@test.com');
    expect(result.total).toBe(0);
    expect(result.byType).toEqual({});
  });
});

describe('CommunicationLogService.logCommunication', () => {
  test('appends a row with escapeForFormula-protected emails', () => {
    var sheet = createMockSheet(SHEETS.COMMUNICATION_LOG, [['ID']]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = CommunicationLogService.logCommunication(
      'steward@test.com', '=EVIL@test.com', 'phone', 'Test', 'Notes'
    );
    expect(result.success).toBe(true);
    expect(result.id).toBeTruthy();
    // Verify escapeForFormula was applied (the mock prepends a tick for formula-like strings)
    var appendedRow = sheet.appendRow.mock.calls[0][0];
    // The member email (index 1) should have been sanitized
    expect(appendedRow[1]).not.toBe('=EVIL@test.com');
  });
});

// ============================================================================
// KnowledgeBaseService
// ============================================================================

describe('KnowledgeBaseService.searchArticles', () => {
  test('returns empty array when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    expect(KnowledgeBaseService.searchArticles('test')).toEqual([]);
  });

  test('returns all articles for empty query', () => {
    var HEADERS = ['ID', 'Article Number', 'Title', 'Summary', 'Full Text', 'Related Grievance Types', 'Tags'];
    var data = [
      HEADERS,
      ['A1', 'Art 5', 'Discipline', 'Summary of Art 5', 'Full text...', 'Discipline', 'contract,rules'],
    ];
    var sheet = createMockSheet(SHEETS.KNOWLEDGE_BASE, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = KnowledgeBaseService.searchArticles('');
    expect(result.length).toBe(1);
  });

  test('filters by query match', () => {
    var HEADERS = ['ID', 'Article Number', 'Title', 'Summary', 'Full Text', 'Related Grievance Types', 'Tags'];
    var data = [
      HEADERS,
      ['A1', 'Art 5', 'Discipline Procedures', 'How to handle discipline', 'Full text...', 'Discipline', 'contract'],
      ['A2', 'Art 12', 'Overtime Rules', 'Overtime calculation', 'Full text...', 'Workload', 'pay'],
    ];
    var sheet = createMockSheet(SHEETS.KNOWLEDGE_BASE, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = KnowledgeBaseService.searchArticles('discipline');
    expect(result.length).toBe(1);
    expect(result[0].title).toBe('Discipline Procedures');
    // Full text should NOT be included in search results
    expect(result[0].fullText).toBeUndefined();
  });
});

describe('KnowledgeBaseService.getArticle', () => {
  test('returns null when article not found', () => {
    var HEADERS = ['ID', 'Article Number', 'Title', 'Summary', 'Full Text', 'Related Grievance Types', 'Tags'];
    var data = [HEADERS, ['A1', 'Art 5', 'Discipline', 'Summary', 'Full text', 'Discipline', 'tags']];
    var sheet = createMockSheet(SHEETS.KNOWLEDGE_BASE, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    expect(KnowledgeBaseService.getArticle('NONEXISTENT')).toBeNull();
  });

  test('returns article with full text when found', () => {
    var HEADERS = ['ID', 'Article Number', 'Title', 'Summary', 'Full Text', 'Related Grievance Types', 'Tags'];
    var data = [HEADERS, ['A1', 'Art 5', 'Discipline', 'Summary', 'Full body text here', 'Discipline', 'tags']];
    var sheet = createMockSheet(SHEETS.KNOWLEDGE_BASE, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = KnowledgeBaseService.getArticle('A1');
    expect(result).not.toBeNull();
    expect(result.fullText).toBe('Full body text here');
  });
});

// ============================================================================
// DigestService
// ============================================================================

describe('DigestService.getPreferences', () => {
  test('returns defaults when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    var result = DigestService.getPreferences('user@test.com');
    expect(result.frequency).toBe('immediate');
    expect(result.types).toBe('all');
  });

  test('returns defaults when user not found in sheet', () => {
    var HEADERS = ['Email', 'Frequency', 'Types', 'Last Sent'];
    var data = [HEADERS, ['other@test.com', 'daily', 'grievances', new Date()]];
    var sheet = createMockSheet(SHEETS.NOTIFICATION_PREFS, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = DigestService.getPreferences('user@test.com');
    expect(result.frequency).toBe('immediate');
  });

  test('returns stored preferences for known user', () => {
    var HEADERS = ['Email', 'Frequency', 'Types', 'Last Sent'];
    var data = [HEADERS, ['user@test.com', 'weekly', 'grievances,meetings', new Date()]];
    var sheet = createMockSheet(SHEETS.NOTIFICATION_PREFS, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = DigestService.getPreferences('user@test.com');
    expect(result.frequency).toBe('weekly');
    expect(result.types).toBe('grievances,meetings');
  });
});

describe('DigestService.setPreferences', () => {
  test('validates frequency to allowed values', () => {
    var HEADERS = ['Email', 'Frequency', 'Types', 'Last Sent'];
    var data = [HEADERS];
    var sheet = createMockSheet(SHEETS.NOTIFICATION_PREFS, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = DigestService.setPreferences('user@test.com', 'invalid-freq', 'all');
    expect(result.success).toBe(true);
    // Should have appended with 'immediate' as the default fallback
    var appendedRow = sheet.appendRow.mock.calls[0][0];
    expect(appendedRow[1]).toBe('immediate');
  });
});

// ============================================================================
// DocumentChecklistService
// ============================================================================

describe('DocumentChecklistService.getChecklist', () => {
  test('returns empty checklist for unknown step', () => {
    var result = DocumentChecklistService.getChecklist('CASE-1', 'Unknown Step');
    expect(result).toEqual([]);
  });

  test('returns required documents for Step I', () => {
    var result = DocumentChecklistService.getChecklist('CASE-1', 'Step I');
    expect(result.length).toBe(2);
    expect(result[0].name).toBe('Grievance Form');
    expect(result[0].required).toBe(true);
    expect(result[0].found).toBe(false);
  });

  test('returns more documents for Arbitration', () => {
    var result = DocumentChecklistService.getChecklist('CASE-1', 'Arbitration');
    expect(result.length).toBe(4);
  });
});

describe('DocumentChecklistService.getCompletionStatus', () => {
  test('returns 0% for unfound documents', () => {
    var result = DocumentChecklistService.getCompletionStatus('CASE-1', 'Step I');
    expect(result.total).toBe(2);
    expect(result.found).toBe(0);
    expect(result.complete).toBe(false);
    expect(result.percentage).toBe(0);
  });

  test('returns 0 total for unknown step', () => {
    var result = DocumentChecklistService.getCompletionStatus('CASE-1', 'Unknown');
    expect(result.total).toBe(0);
    expect(result.percentage).toBe(0);
  });
});

// ============================================================================
// EscalationEngine
// ============================================================================

describe('EscalationEngine.getRecommendation', () => {
  test('returns null when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    expect(EscalationEngine.getRecommendation('CASE-1')).toBeNull();
  });

  test('returns null when case not found', () => {
    var HEADERS = ['Grievance ID', 'Status', 'Current Step', 'Next Deadline', 'Issue Category', 'Unit'];
    var data = [HEADERS, ['CASE-99', 'Open', 'Step I', new Date(), 'Discipline', 'Unit A']];
    var sheet = createMockSheet(SHEETS.GRIEVANCE_LOG, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    expect(EscalationEngine.getRecommendation('CASE-1')).toBeNull();
  });

  test('recommends escalation for denied case', () => {
    var HEADERS = ['Grievance ID', 'Status', 'Current Step', 'Next Deadline', 'Issue Category', 'Unit'];
    var data = [
      HEADERS,
      ['CASE-1', 'Denied', 'Step I', new Date(), 'Discipline', 'Unit A'],
    ];
    var sheet = createMockSheet(SHEETS.GRIEVANCE_LOG, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = EscalationEngine.getRecommendation('CASE-1');
    expect(result).not.toBeNull();
    expect(result.shouldEscalate).toBe(true);
    expect(result.confidence).toBe('high');
    expect(result.suggestedStep).toBe('Step II');
  });

  test('recommends escalation for overdue deadline', () => {
    var overdue = new Date(Date.now() - 5 * 86400000); // 5 days ago
    var HEADERS = ['Grievance ID', 'Status', 'Current Step', 'Next Deadline', 'Issue Category', 'Unit'];
    var data = [
      HEADERS,
      ['CASE-2', 'Open', 'Step II', overdue, 'Workload', 'Unit B'],
    ];
    var sheet = createMockSheet(SHEETS.GRIEVANCE_LOG, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = EscalationEngine.getRecommendation('CASE-2');
    expect(result.shouldEscalate).toBe(true);
    expect(result.confidence).toBe('high');
    expect(result.reasons.some(function(r) { return r.indexOf('overdue') !== -1; })).toBe(true);
  });
});

// ============================================================================
// ReportService
// ============================================================================

describe('ReportService.generateMonthlyReport', () => {
  test('returns empty report when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    var result = ReportService.generateMonthlyReport();
    expect(result.success).toBe(false);
  });

  test('returns empty summary when no grievance sheet', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    var result = ReportService.generateMonthlyReport();
    expect(result.success).toBe(true);
    expect(result.report.summary.totalCases).toBe(0);
  });
});

describe('ReportService.generateReportHtml', () => {
  test('returns fallback when report is null', () => {
    var result = ReportService.generateReportHtml(null);
    expect(result).toContain('No report data');
  });

  test('returns HTML with report data', () => {
    var report = {
      title: 'Monthly Summary',
      period: '03/01/2026 - 03/31/2026',
      summary: { totalCases: 10, newCases: 3, resolved: 5, pending: 2, overdue: 1 },
      byCategory: { 'Discipline': 4 },
      bySteward: {},
      byUnit: {},
    };
    var html = ReportService.generateReportHtml(report);
    expect(html).toContain('Monthly Summary');
    expect(html).toContain('10');
  });
});

// ============================================================================
// TwoFactorService
// ============================================================================

describe('TwoFactorService.generateCode', () => {
  test('rejects empty email', () => {
    var result = TwoFactorService.generateCode('');
    expect(result.success).toBe(false);
    expect(result.error).toContain('Email required');
  });

  test('sends code for valid email', () => {
    var result = TwoFactorService.generateCode('user@test.com');
    expect(result.success).toBe(true);
    expect(result.message).toContain('Verification code sent');
  });
});

describe('TwoFactorService.verifyCode', () => {
  test('rejects empty inputs', () => {
    var result = TwoFactorService.verifyCode('', '');
    expect(result.success).toBe(false);
    expect(result.error).toContain('required');
  });

  test('rejects expired or non-existent code', () => {
    var result = TwoFactorService.verifyCode('user@test.com', '123456');
    expect(result.success).toBe(false);
    expect(result.error).toContain('expired');
  });
});

describe('TwoFactorService.hasValidSession', () => {
  test('returns false for null email', () => {
    expect(TwoFactorService.hasValidSession(null)).toBe(false);
  });

  test('returns false when no session exists', () => {
    expect(TwoFactorService.hasValidSession('user@test.com')).toBe(false);
  });
});

// ============================================================================
// SMSService
// ============================================================================

describe('SMSService.configureProvider', () => {
  test('rejects missing fields', () => {
    var result = SMSService.configureProvider('', '', '');
    expect(result.success).toBe(false);
    expect(result.message).toContain('required');
  });

  test('rejects invalid phone format', () => {
    var result = SMSService.configureProvider('ACXXX', 'token123', '5551234567');
    expect(result.success).toBe(false);
    expect(result.message).toContain('E.164');
  });

  test('accepts valid E.164 number', () => {
    var result = SMSService.configureProvider('ACXXX', 'token123', '+15551234567');
    expect(result.success).toBe(true);
  });
});

describe('SMSService.isConfigured', () => {
  test('returns false when no credentials set', () => {
    expect(SMSService.isConfigured()).toBe(false);
  });

  test('returns true when all credentials set', () => {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('TWILIO_ACCOUNT_SID', 'ACxxxx');
    props.setProperty('TWILIO_AUTH_TOKEN', 'tokenxxx');
    props.setProperty('TWILIO_FROM_NUMBER', '+15551234567');
    expect(SMSService.isConfigured()).toBe(true);
  });
});

describe('SMSService.sendSMS', () => {
  test('rejects missing recipient', () => {
    var result = SMSService.sendSMS('', 'Hello');
    expect(result.success).toBe(false);
    expect(result.error).toContain('required');
  });

  test('rejects missing message', () => {
    var result = SMSService.sendSMS('user@test.com', '');
    expect(result.success).toBe(false);
    expect(result.error).toContain('required');
  });
});

// ============================================================================
// Global wrapper auth gating
// ============================================================================

describe('NewFeatureServices global wrappers — auth gating', () => {
  test('steward wrappers reject null auth', () => {
    _requireStewardAuth.mockReturnValue(null);

    expect(dataGetHandoffNotes('bad', 'CASE-1')).toEqual([]);
    expect(dataAddHandoffNote('bad', 'CASE-1', 'note').success).toBe(false);
    expect(dataGetMentorshipPairings('bad')).toEqual([]);
    expect(dataLogCommunication('bad', 'e', 't', 's', 'n').success).toBe(false);
    expect(dataGetCommunicationLog('bad', 'e')).toEqual([]);
    expect(dataGetDocumentChecklist('bad', 'c', 's', 'f')).toEqual([]);
    expect(dataGetEscalationRecommendation('bad', 'c')).toBeNull();
    expect(dataGenerateMonthlyReport('bad').success).toBe(false);

    // Restore
    _requireStewardAuth.mockReturnValue('steward@example.com');
  });

  test('caller wrappers reject null auth', () => {
    _resolveCallerEmail.mockReturnValue(null);

    expect(dataSearchKnowledgeBase('bad', 'q')).toEqual([]);
    expect(dataGetKnowledgeBaseArticle('bad', 'id')).toBeNull();
    expect(dataGetDigestPreferences('bad').frequency).toBe('immediate');
    expect(dataSetDigestPreferences('bad', 'daily', 'all').success).toBe(false);
    expect(dataRequest2FACode('bad').success).toBe(false);
    expect(dataVerify2FACode('bad', '123').success).toBe(false);

    // Restore
    _resolveCallerEmail.mockReturnValue('test@example.com');
  });
});
