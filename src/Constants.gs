/**
 * 509 Dashboard - Constants and Configuration
 *
 * Single source of truth for all configuration constants.
 * This file must be loaded first in the build order.
 *
 * @version 2.2.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// VERSION INFO
// ============================================================================

/**
 * Version information for build system and display
 * @const {Object}
 */
var VERSION_INFO = {
  MAJOR: 2,
  MINOR: 2,
  PATCH: 0,
  BUILD: 'v3.48',
  CODENAME: 'Survey Verification & Data Integrity'
};

// ============================================================================
// SHEET NAMES
// ============================================================================

/**
 * Sheet name constants - use these instead of hardcoded strings
 * @const {Object}
 */
var SHEETS = {
  CONFIG: 'Config',
  MEMBER_DIR: 'Member Directory',
  GRIEVANCE_LOG: 'Grievance Log',
  // Dashboard sheets
  DASHBOARD: '💼 Dashboard',
  INTERACTIVE: '🎯 Custom View',
  // Hidden calculation sheets (self-healing formulas)
  GRIEVANCE_CALC: '_Grievance_Calc',
  GRIEVANCE_FORMULAS: '_Grievance_Formulas',
  MEMBER_LOOKUP: '_Member_Lookup',
  STEWARD_CONTACT_CALC: '_Steward_Contact_Calc',
  DASHBOARD_CALC: '_Dashboard_Calc',
  STEWARD_PERFORMANCE_CALC: '_Steward_Performance_Calc',
  // Optional source sheets
  MEETING_ATTENDANCE: '📅 Meeting Attendance',
  VOLUNTEER_HOURS: '🤝 Volunteer Hours',
  // Test Results
  TEST_RESULTS: 'Test Results',
  // Function Checklist
  FUNCTION_CHECKLIST: 'Function Checklist',
  // Audit Log (hidden)
  AUDIT_LOG: '_Audit_Log',
  // Satisfaction & Feedback sheets
  SATISFACTION: '📊 Member Satisfaction',
  FEEDBACK: '💡 Feedback & Development',
  // Help & Documentation sheets
  GETTING_STARTED: '📚 Getting Started',
  FAQ: '❓ FAQ',
  CONFIG_GUIDE: '📖 Config Guide'
};

// ============================================================================
// COLOR SCHEME
// ============================================================================

/**
 * Color scheme constants for consistent branding
 * @const {Object}
 */
var COLORS = {
  PRIMARY_PURPLE: '#7C3AED',    // Main brand purple
  UNION_GREEN: '#059669',       // Union/success green
  SOLIDARITY_RED: '#DC2626',    // Alert/urgent red
  PRIMARY_BLUE: '#7EC8E3',      // Light blue
  ACCENT_ORANGE: '#F97316',     // Warnings/attention
  LIGHT_GRAY: '#F3F4F6',        // Backgrounds
  TEXT_DARK: '#1F2937',         // Primary text
  WHITE: '#FFFFFF',             // White
  HEADER_BG: '#7C3AED',         // Header background (same as primary)
  HEADER_TEXT: '#FFFFFF'        // Header text color
};

// ============================================================================
// MEMBER DIRECTORY COLUMNS (32 columns total: A-AF)
// ============================================================================

/**
 * Member Directory column positions (1-indexed)
 * CRITICAL: ALL column references must use these constants
 * @const {Object}
 */
var MEMBER_COLS = {
  // Section 1: Identity & Core Info (A-D)
  MEMBER_ID: 1,                    // A
  FIRST_NAME: 2,                   // B
  LAST_NAME: 3,                    // C
  JOB_TITLE: 4,                    // D

  // Section 2: Location & Work (E-G)
  WORK_LOCATION: 5,                // E
  UNIT: 6,                         // F
  OFFICE_DAYS: 7,                  // G - Multi-select: days member works in office

  // Section 3: Contact Information (H-K)
  EMAIL: 8,                        // H
  PHONE: 9,                        // I
  PREFERRED_COMM: 10,              // J - Multi-select: preferred communication methods
  BEST_TIME: 11,                   // K - Multi-select: best times to reach member

  // Section 4: Organizational Structure (L-P)
  SUPERVISOR: 12,                  // L
  MANAGER: 13,                     // M
  IS_STEWARD: 14,                  // N
  COMMITTEES: 15,                  // O - Multi-select: which committees steward is in
  ASSIGNED_STEWARD: 16,            // P - Multi-select: assigned steward(s)

  // Section 5: Engagement Metrics (Q-T) - Hidden by default
  LAST_VIRTUAL_MTG: 17,            // Q
  LAST_INPERSON_MTG: 18,           // R
  OPEN_RATE: 19,                   // S
  VOLUNTEER_HOURS: 20,             // T

  // Section 6: Member Interests (U-X) - Hidden by default
  INTEREST_LOCAL: 21,              // U
  INTEREST_CHAPTER: 22,            // V
  INTEREST_ALLIED: 23,             // W
  HOME_TOWN: 24,                   // X - Connection building

  // Section 7: Steward Contact Tracking (Y-AA)
  RECENT_CONTACT_DATE: 25,         // Y
  CONTACT_STEWARD: 26,             // Z
  CONTACT_NOTES: 27,               // AA

  // Section 8: Grievance Management (AB-AE)
  HAS_OPEN_GRIEVANCE: 28,          // AB - Script-calculated (static value)
  GRIEVANCE_STATUS: 29,            // AC - Script-calculated (static value)
  NEXT_DEADLINE: 30,               // AD - Script-calculated (static value)
  START_GRIEVANCE: 31,             // AE - Checkbox to start grievance

  // Section 9: Quick Actions (AF)
  QUICK_ACTIONS: 32,               // AF - Checkbox to open Quick Actions dialog

  // ALIASES - For backward compatibility
  LOCATION: 5,                     // Alias for WORK_LOCATION
  DAYS_TO_DEADLINE: 30             // Alias for NEXT_DEADLINE
};

// ============================================================================
// GRIEVANCE LOG COLUMNS (35 columns total: A-AI)
// ============================================================================

/**
 * Grievance Log column positions (1-indexed)
 * CRITICAL: ALL column references must use these constants
 * @const {Object}
 */
var GRIEVANCE_COLS = {
  // Section 1: Identity (A-D)
  GRIEVANCE_ID: 1,        // A - Grievance ID
  MEMBER_ID: 2,           // B - Member ID
  FIRST_NAME: 3,          // C - First Name
  LAST_NAME: 4,           // D - Last Name

  // Section 2: Status & Assignment (E-F)
  STATUS: 5,              // E - Status
  CURRENT_STEP: 6,        // F - Current Step

  // Section 3: Timeline - Filing (G-I)
  INCIDENT_DATE: 7,       // G - Incident Date
  FILING_DEADLINE: 8,     // H - Filing Deadline (21d) (auto-calc)
  DATE_FILED: 9,          // I - Date Filed (Step I)

  // Section 4: Timeline - Step I (J-K)
  STEP1_DUE: 10,          // J - Step I Decision Due (30d) (auto-calc)
  STEP1_RCVD: 11,         // K - Step I Decision Rcvd

  // Section 5: Timeline - Step II (L-O)
  STEP2_APPEAL_DUE: 12,   // L - Step II Appeal Due (10d) (auto-calc)
  STEP2_APPEAL_FILED: 13, // M - Step II Appeal Filed
  STEP2_DUE: 14,          // N - Step II Decision Due (30d) (auto-calc)
  STEP2_RCVD: 15,         // O - Step II Decision Rcvd

  // Section 6: Timeline - Step III (P-R)
  STEP3_APPEAL_DUE: 16,   // P - Step III Appeal Due (30d) (auto-calc)
  STEP3_APPEAL_FILED: 17, // Q - Step III Appeal Filed
  DATE_CLOSED: 18,        // R - Date Closed

  // Section 7: Calculated Metrics (S-U)
  DAYS_OPEN: 19,          // S - Days Open (auto-calc)
  NEXT_ACTION_DUE: 20,    // T - Next Action Due (auto-calc)
  DAYS_TO_DEADLINE: 21,   // U - Days to Deadline (auto-calc)

  // Section 8: Case Details (V-W)
  ARTICLES: 22,           // V - Articles Violated
  ISSUE_CATEGORY: 23,     // W - Issue Category

  // Section 9: Contact & Location (X-AA)
  MEMBER_EMAIL: 24,       // X - Member Email
  UNIT: 25,               // Y - Unit
  LOCATION: 26,           // Z - Work Location (Site)
  STEWARD: 27,            // AA - Assigned Steward (Name)

  // Section 10: Resolution (AB)
  RESOLUTION: 28,         // AB - Resolution Summary

  // Section 11: Coordinator Notifications (AC-AF)
  MESSAGE_ALERT: 29,      // AC - Message Alert checkbox
  COORDINATOR_MESSAGE: 30,// AD - Coordinator's message text
  ACKNOWLEDGED_BY: 31,    // AE - Steward who acknowledged
  ACKNOWLEDGED_DATE: 32,  // AF - When steward acknowledged

  // Section 12: Drive Integration (AG-AH)
  DRIVE_FOLDER_ID: 33,    // AG - Google Drive folder ID
  DRIVE_FOLDER_URL: 34,   // AH - Google Drive folder URL

  // Section 13: Quick Actions (AI)
  QUICK_ACTIONS: 35       // AI - Checkbox to open Quick Actions dialog
};

// ============================================================================
// CONFIG COLUMN MAPPING
// ============================================================================

/**
 * Config sheet column positions for dropdown sources
 * @const {Object}
 */
var CONFIG_COLS = {
  // ── EMPLOYMENT INFO ── (A-E)
  JOB_TITLES: 1,              // A
  OFFICE_LOCATIONS: 2,        // B
  UNITS: 3,                   // C
  OFFICE_DAYS: 4,             // D
  YES_NO: 5,                  // E

  // ── SUPERVISION ── (F-G)
  SUPERVISORS: 6,             // F
  MANAGERS: 7,                // G

  // ── STEWARD INFO ── (H-I)
  STEWARDS: 8,                // H
  STEWARD_COMMITTEES: 9,      // I

  // ── GRIEVANCE SETTINGS ── (J-M)
  GRIEVANCE_STATUS: 10,       // J
  GRIEVANCE_STEP: 11,         // K
  ISSUE_CATEGORY: 12,         // L
  ARTICLES: 13,               // M

  // ── LINKS & COORDINATORS ── (N-Q)
  COMM_METHODS: 14,           // N
  GRIEVANCE_COORDINATORS: 15, // O
  GRIEVANCE_FORM_URL: 16,     // P
  CONTACT_FORM_URL: 17,       // Q

  // ── NOTIFICATIONS ── (R-S)
  ADMIN_EMAILS: 18,           // R
  ALERT_DAYS: 19,             // S
  NOTIFICATION_RECIPIENTS: 20, // T

  // ── ORGANIZATION ── (U-X)
  ORG_NAME: 21,               // U
  LOCAL_NUMBER: 22,           // V
  MAIN_ADDRESS: 23,           // W
  MAIN_PHONE: 24,             // X

  // ── INTEGRATION ── (Y-Z)
  DRIVE_FOLDER_ID: 25,        // Y
  CALENDAR_ID: 26,            // Z

  // ── DEADLINES ── (AA-AD)
  FILING_DEADLINE_DAYS: 27,   // AA
  STEP1_RESPONSE_DAYS: 28,    // AB
  STEP2_APPEAL_DAYS: 29,      // AC
  STEP2_RESPONSE_DAYS: 30,    // AD

  // ── MULTI-SELECT OPTIONS ── (AE-AF)
  BEST_TIMES: 31,             // AE
  HOME_TOWNS: 32,             // AF

  // ── CONTRACT & LEGAL ── (AG-AJ)
  CONTRACT_GRIEVANCE: 33,     // AG
  CONTRACT_DISCIPLINE: 34,    // AH
  CONTRACT_WORKLOAD: 35,      // AI
  CONTRACT_NAME: 36,          // AJ

  // ── ORG IDENTITY ── (AK-AM)
  UNION_PARENT: 37,           // AK
  STATE_REGION: 38,           // AL
  ORG_WEBSITE: 39,            // AM

  // ── EXTENDED CONTACT ── (AN-AQ)
  OFFICE_ADDRESSES: 40,       // AN
  MAIN_FAX: 41,               // AO
  MAIN_CONTACT_NAME: 42,      // AP
  MAIN_CONTACT_EMAIL: 43,     // AQ

  // ── FORM LINKS ── (AR)
  SATISFACTION_FORM_URL: 44   // AR - Member Satisfaction Survey form URL
};

// ============================================================================
// SATISFACTION SURVEY COLUMNS (Google Form Response + Summary)
// ============================================================================

/**
 * Member Satisfaction Survey column positions (1-indexed)
 *
 * FORM RESPONSE AREA (A-BQ): Auto-populated by Google Form link
 * - Column A: Timestamp (auto-generated by Google Forms)
 * - Columns B-BQ: 68 question responses
 *
 * SUMMARY/CHART DATA AREA (Column BT onwards): Aggregated metrics for charts
 *
 * @const {Object}
 */
var SATISFACTION_COLS = {
  // ── FORM RESPONSE COLUMNS (Auto-created by Google Form) ──
  TIMESTAMP: 1,                   // A - Auto-generated by Google Forms

  // Work Context (Q1-5)
  Q1_WORKSITE: 2,                 // B
  Q2_ROLE: 3,                     // C
  Q3_SHIFT: 4,                    // D
  Q4_TIME_IN_ROLE: 5,             // E
  Q5_STEWARD_CONTACT: 6,          // F (branching: Yes → 3A, No → 3B)

  // Overall Satisfaction (Q6-9) - Scale 1-10
  Q6_SATISFIED_REP: 7,            // G
  Q7_TRUST_UNION: 8,              // H
  Q8_FEEL_PROTECTED: 9,           // I
  Q9_RECOMMEND: 10,               // J

  // Steward Ratings 3A (Q10-17) - For those with steward contact
  Q10_TIMELY_RESPONSE: 11,        // K
  Q11_TREATED_RESPECT: 12,        // L
  Q12_EXPLAINED_OPTIONS: 13,      // M
  Q13_FOLLOWED_THROUGH: 14,       // N
  Q14_ADVOCATED: 15,              // O
  Q15_SAFE_CONCERNS: 16,          // P
  Q16_CONFIDENTIALITY: 17,        // Q
  Q17_STEWARD_IMPROVE: 18,        // R (paragraph)

  // Steward Access 3B (Q18-20) - For those without steward contact
  Q18_KNOW_CONTACT: 19,           // S
  Q19_CONFIDENT_HELP: 20,         // T
  Q20_EASY_FIND: 21,              // U

  // Chapter Effectiveness (Q21-25)
  Q21_UNDERSTAND_ISSUES: 22,      // V
  Q22_CHAPTER_COMM: 23,           // W
  Q23_ORGANIZES: 24,              // X
  Q24_REACH_CHAPTER: 25,          // Y
  Q25_FAIR_REP: 26,               // Z

  // Local Leadership (Q26-31)
  Q26_DECISIONS_CLEAR: 27,        // AA
  Q27_UNDERSTAND_PROCESS: 28,     // AB
  Q28_TRANSPARENT_FINANCE: 29,    // AC
  Q29_ACCOUNTABLE: 30,            // AD
  Q30_FAIR_PROCESSES: 31,         // AE
  Q31_WELCOMES_OPINIONS: 32,      // AF

  // Contract Enforcement (Q32-36)
  Q32_ENFORCES_CONTRACT: 33,      // AG
  Q33_REALISTIC_TIMELINES: 34,    // AH
  Q34_CLEAR_UPDATES: 35,          // AI
  Q35_FRONTLINE_PRIORITY: 36,     // AJ
  Q36_FILED_GRIEVANCE: 37,        // AK (branching: Yes → 6A, No → 7)

  // Representation Process 6A (Q37-40) - For those who filed grievance
  Q37_UNDERSTOOD_STEPS: 38,       // AL
  Q38_FELT_SUPPORTED: 39,         // AM
  Q39_UPDATES_OFTEN: 40,          // AN
  Q40_OUTCOME_JUSTIFIED: 41,      // AO

  // Communication Quality (Q41-45)
  Q41_CLEAR_ACTIONABLE: 42,       // AP
  Q42_ENOUGH_INFO: 43,            // AQ
  Q43_FIND_EASILY: 44,            // AR
  Q44_ALL_SHIFTS: 45,             // AS
  Q45_MEETINGS_WORTH: 46,         // AT

  // Member Voice & Culture (Q46-50)
  Q46_VOICE_MATTERS: 47,          // AU
  Q47_SEEKS_INPUT: 48,            // AV
  Q48_DIGNITY: 49,                // AW
  Q49_NEWER_SUPPORTED: 50,        // AX
  Q50_CONFLICT_RESPECT: 51,       // AY

  // Value & Collective Action (Q51-55)
  Q51_GOOD_VALUE: 52,             // AZ
  Q52_PRIORITIES_NEEDS: 53,       // BA
  Q53_PREPARED_MOBILIZE: 54,      // BB
  Q54_HOW_INVOLVED: 55,           // BC
  Q55_WIN_TOGETHER: 56,           // BD

  // Scheduling/Office Days (Q56-63)
  Q56_UNDERSTAND_CHANGES: 57,     // BE
  Q57_ADEQUATELY_INFORMED: 58,    // BF
  Q58_CLEAR_CRITERIA: 59,         // BG
  Q59_WORK_EXPECTATIONS: 60,      // BH
  Q60_EFFECTIVE_OUTCOMES: 61,     // BI
  Q61_SUPPORTS_WELLBEING: 62,     // BJ
  Q62_CONCERNS_SERIOUS: 63,       // BK
  Q63_SCHEDULING_CHALLENGE: 64,   // BL (paragraph)

  // Priorities & Close (Q64-68)
  Q64_TOP_PRIORITIES: 65,         // BM (checkboxes - comma separated)
  Q65_ONE_CHANGE: 66,             // BN (paragraph)
  Q66_KEEP_DOING: 67,             // BO (paragraph)
  Q67_ADDITIONAL: 68,             // BP (paragraph)
  Q68_SUBMIT: 69,                 // BQ (if present)

  // ── SUMMARY/CHART DATA AREA (Column BT onwards) ──
  SUMMARY_START: 72,              // BT - Start of summary section

  // Section averages for charts
  AVG_OVERALL_SAT: 72,            // BT - Avg of Q6-Q9
  AVG_STEWARD_RATING: 73,         // BU - Avg of Q10-Q16
  AVG_STEWARD_ACCESS: 74,         // BV - Avg of Q18-Q20
  AVG_CHAPTER: 75,                // BW - Avg of Q21-Q25
  AVG_LEADERSHIP: 76,             // BX - Avg of Q26-Q31
  AVG_CONTRACT: 77,               // BY - Avg of Q32-Q35
  AVG_REPRESENTATION: 78,         // BZ - Avg of Q37-Q40
  AVG_COMMUNICATION: 79,          // CA - Avg of Q41-Q45
  AVG_MEMBER_VOICE: 80,           // CB - Avg of Q46-Q50
  AVG_VALUE_ACTION: 81,           // CC - Avg of Q51-Q55
  AVG_SCHEDULING: 82,             // CD - Avg of Q56-Q62

  // ── VERIFICATION & TRACKING COLUMNS (CE onwards) ──
  EMAIL: 83,                      // CE - Email address from form submission
  VERIFIED: 84,                   // CF - Yes / Pending Review / Rejected
  MATCHED_MEMBER_ID: 85,          // CG - Member ID if email matched
  QUARTER: 86,                    // CH - Quarter string (e.g., "2026-Q1")
  IS_LATEST: 87,                  // CI - Yes/No - Is this the latest for this member this quarter?
  SUPERSEDED_BY: 88,              // CJ - Row number of newer response (if superseded)
  REVIEWER_NOTES: 89              // CK - Notes from reviewer
};

/**
 * Survey section definitions for grouping and analysis
 * @const {Object}
 */
var SATISFACTION_SECTIONS = {
  WORK_CONTEXT: { name: 'Work Context', questions: [2,3,4,5,6], scale: false },
  OVERALL_SAT: { name: 'Overall Satisfaction', questions: [7,8,9,10], scale: true },
  STEWARD_3A: { name: 'Steward Ratings', questions: [11,12,13,14,15,16,17], scale: true },
  STEWARD_3B: { name: 'Steward Access', questions: [19,20,21], scale: true },
  CHAPTER: { name: 'Chapter Effectiveness', questions: [22,23,24,25,26], scale: true },
  LEADERSHIP: { name: 'Local Leadership', questions: [27,28,29,30,31,32], scale: true },
  CONTRACT: { name: 'Contract Enforcement', questions: [33,34,35,36], scale: true },
  REPRESENTATION: { name: 'Representation Process', questions: [38,39,40,41], scale: true },
  COMMUNICATION: { name: 'Communication Quality', questions: [42,43,44,45,46], scale: true },
  MEMBER_VOICE: { name: 'Member Voice & Culture', questions: [47,48,49,50,51], scale: true },
  VALUE_ACTION: { name: 'Value & Collective Action', questions: [52,53,54,55,56], scale: true },
  SCHEDULING: { name: 'Scheduling/Office Days', questions: [57,58,59,60,61,62,63], scale: true },
  PRIORITIES: { name: 'Priorities & Close', questions: [65,66,67,68], scale: false }
};

// ============================================================================
// FEEDBACK & DEVELOPMENT COLUMNS (11 columns: A-K)
// ============================================================================

/**
 * Feedback & Development column positions (1-indexed)
 * @const {Object}
 */
var FEEDBACK_COLS = {
  TIMESTAMP: 1,                // A - Auto-generated timestamp
  SUBMITTED_BY: 2,             // B - Who submitted the feedback
  CATEGORY: 3,                 // C - Area of the system
  TYPE: 4,                     // D - Bug, Feature Request, Improvement
  PRIORITY: 5,                 // E - Low, Medium, High, Critical
  TITLE: 6,                    // F - Short title
  DESCRIPTION: 7,              // G - Detailed description
  STATUS: 8,                   // H - New, In Progress, Resolved, Won't Fix
  ASSIGNED_TO: 9,              // I - Who is working on it
  RESOLUTION: 10,              // J - How it was resolved
  NOTES: 11                    // K - Additional notes
};

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Convert column number to letter notation (e.g., 1 -> A, 27 -> AA)
 * @param {number} columnNumber - Column number (1-indexed)
 * @returns {string} Column letter(s)
 */
function getColumnLetter(columnNumber) {
  var letter = '';
  while (columnNumber > 0) {
    var remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

/**
 * Convert column letter to number (e.g., A -> 1, AA -> 27)
 * @param {string} columnLetter - Column letter(s)
 * @returns {number} Column number (1-indexed)
 */
function getColumnNumber(columnLetter) {
  var result = 0;
  for (var i = 0; i < columnLetter.length; i++) {
    result = result * 26 + (columnLetter.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Map a Member Directory row array to a structured object
 * @param {Array} row - Row data array from Member Directory
 * @returns {Object} Structured member object
 */
function mapMemberRow(row) {
  return {
    memberId: row[MEMBER_COLS.MEMBER_ID - 1] || '',
    firstName: row[MEMBER_COLS.FIRST_NAME - 1] || '',
    lastName: row[MEMBER_COLS.LAST_NAME - 1] || '',
    fullName: (row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (row[MEMBER_COLS.LAST_NAME - 1] || ''),
    jobTitle: row[MEMBER_COLS.JOB_TITLE - 1] || '',
    workLocation: row[MEMBER_COLS.WORK_LOCATION - 1] || '',
    unit: row[MEMBER_COLS.UNIT - 1] || '',
    officeDays: row[MEMBER_COLS.OFFICE_DAYS - 1] || '',
    email: row[MEMBER_COLS.EMAIL - 1] || '',
    phone: row[MEMBER_COLS.PHONE - 1] || '',
    preferredComm: row[MEMBER_COLS.PREFERRED_COMM - 1] || '',
    bestTime: row[MEMBER_COLS.BEST_TIME - 1] || '',
    supervisor: row[MEMBER_COLS.SUPERVISOR - 1] || '',
    manager: row[MEMBER_COLS.MANAGER - 1] || '',
    isSteward: row[MEMBER_COLS.IS_STEWARD - 1] || '',
    committees: row[MEMBER_COLS.COMMITTEES - 1] || '',
    assignedSteward: row[MEMBER_COLS.ASSIGNED_STEWARD - 1] || '',
    lastVirtualMtg: row[MEMBER_COLS.LAST_VIRTUAL_MTG - 1] || '',
    lastInPersonMtg: row[MEMBER_COLS.LAST_INPERSON_MTG - 1] || '',
    openRate: row[MEMBER_COLS.OPEN_RATE - 1] || '',
    volunteerHours: row[MEMBER_COLS.VOLUNTEER_HOURS - 1] || '',
    interestLocal: row[MEMBER_COLS.INTEREST_LOCAL - 1] || '',
    interestChapter: row[MEMBER_COLS.INTEREST_CHAPTER - 1] || '',
    interestAllied: row[MEMBER_COLS.INTEREST_ALLIED - 1] || '',
    homeTown: row[MEMBER_COLS.HOME_TOWN - 1] || '',
    recentContactDate: row[MEMBER_COLS.RECENT_CONTACT_DATE - 1] || '',
    contactSteward: row[MEMBER_COLS.CONTACT_STEWARD - 1] || '',
    contactNotes: row[MEMBER_COLS.CONTACT_NOTES - 1] || '',
    hasOpenGrievance: row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1] || '',
    grievanceStatus: row[MEMBER_COLS.GRIEVANCE_STATUS - 1] || '',
    nextDeadline: row[MEMBER_COLS.NEXT_DEADLINE - 1] || '',
    startGrievance: row[MEMBER_COLS.START_GRIEVANCE - 1] || false
  };
}

/**
 * Map a Grievance Log row array to a structured object
 * @param {Array} row - Row data array from Grievance Log
 * @returns {Object} Structured grievance object
 */
function mapGrievanceRow(row) {
  return {
    grievanceId: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '',
    memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1] || '',
    firstName: row[GRIEVANCE_COLS.FIRST_NAME - 1] || '',
    lastName: row[GRIEVANCE_COLS.LAST_NAME - 1] || '',
    fullName: (row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || ''),
    status: row[GRIEVANCE_COLS.STATUS - 1] || '',
    currentStep: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || '',
    incidentDate: row[GRIEVANCE_COLS.INCIDENT_DATE - 1] || '',
    filingDeadline: row[GRIEVANCE_COLS.FILING_DEADLINE - 1] || '',
    dateFiled: row[GRIEVANCE_COLS.DATE_FILED - 1] || '',
    step1Due: row[GRIEVANCE_COLS.STEP1_DUE - 1] || '',
    step1Rcvd: row[GRIEVANCE_COLS.STEP1_RCVD - 1] || '',
    step2AppealDue: row[GRIEVANCE_COLS.STEP2_APPEAL_DUE - 1] || '',
    step2AppealFiled: row[GRIEVANCE_COLS.STEP2_APPEAL_FILED - 1] || '',
    step2Due: row[GRIEVANCE_COLS.STEP2_DUE - 1] || '',
    step2Rcvd: row[GRIEVANCE_COLS.STEP2_RCVD - 1] || '',
    step3AppealDue: row[GRIEVANCE_COLS.STEP3_APPEAL_DUE - 1] || '',
    step3AppealFiled: row[GRIEVANCE_COLS.STEP3_APPEAL_FILED - 1] || '',
    dateClosed: row[GRIEVANCE_COLS.DATE_CLOSED - 1] || '',
    daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || '',
    nextActionDue: row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1] || '',
    daysToDeadline: row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] || '',
    articles: row[GRIEVANCE_COLS.ARTICLES - 1] || '',
    issueCategory: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '',
    memberEmail: row[GRIEVANCE_COLS.MEMBER_EMAIL - 1] || '',
    unit: row[GRIEVANCE_COLS.UNIT - 1] || '',
    location: row[GRIEVANCE_COLS.LOCATION - 1] || '',
    steward: row[GRIEVANCE_COLS.STEWARD - 1] || '',
    resolution: row[GRIEVANCE_COLS.RESOLUTION - 1] || '',
    messageAlert: row[GRIEVANCE_COLS.MESSAGE_ALERT - 1] || false,
    coordinatorMessage: row[GRIEVANCE_COLS.COORDINATOR_MESSAGE - 1] || '',
    acknowledgedBy: row[GRIEVANCE_COLS.ACKNOWLEDGED_BY - 1] || '',
    acknowledgedDate: row[GRIEVANCE_COLS.ACKNOWLEDGED_DATE - 1] || '',
    driveFolderId: row[GRIEVANCE_COLS.DRIVE_FOLDER_ID - 1] || '',
    driveFolderUrl: row[GRIEVANCE_COLS.DRIVE_FOLDER_URL - 1] || ''
  };
}

/**
 * Get all member header labels in order
 * @returns {Array} Array of header labels for Member Directory
 */
function getMemberHeaders() {
  return [
    'Member ID', 'First Name', 'Last Name', 'Job Title',
    'Work Location', 'Unit', 'Office Days',
    'Email', 'Phone', 'Preferred Communication', 'Best Time to Contact',
    'Supervisor', 'Manager', 'Is Steward', 'Committees', 'Assigned Steward',
    'Last Virtual Mtg', 'Last In-Person Mtg', 'Open Rate %', 'Volunteer Hours',
    'Interest: Local', 'Interest: Chapter', 'Interest: Allied', 'Home Town',
    'Recent Contact Date', 'Contact Steward', 'Contact Notes',
    'Has Open Grievance?', 'Grievance Status', 'Days to Deadline', 'Start Grievance',
    '⚡ Actions'
  ];
}

/**
 * Get all grievance header labels in order
 * @returns {Array} Array of header labels for Grievance Log
 */
function getGrievanceHeaders() {
  return [
    'Grievance ID', 'Member ID', 'First Name', 'Last Name',
    'Status', 'Current Step',
    'Incident Date', 'Filing Deadline', 'Date Filed',
    'Step I Due', 'Step I Rcvd',
    'Step II Appeal Due', 'Step II Appeal Filed', 'Step II Due', 'Step II Rcvd',
    'Step III Appeal Due', 'Step III Appeal Filed', 'Date Closed',
    'Days Open', 'Next Action Due', 'Days to Deadline',
    'Articles Violated', 'Issue Category',
    'Member Email', 'Unit', 'Work Location', 'Assigned Steward',
    'Resolution',
    'Message Alert', 'Coordinator Message', 'Acknowledged By', 'Acknowledged Date',
    'Drive Folder ID', 'Drive Folder URL',
    '⚡ Actions'
  ];
}

// ============================================================================
// VALIDATION VALUES
// ============================================================================

/**
 * Default values for Config sheet dropdowns
 */
var DEFAULT_CONFIG = {
  OFFICE_DAYS: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
  YES_NO: ['Yes', 'No'],
  // Status includes both workflow states (Open, Pending, In Arbitration) AND outcomes (Won, Denied, Settled, Withdrawn)
  // This single-column design allows Dashboard metrics to count outcomes directly from STATUS column
  GRIEVANCE_STATUS: ['Open', 'Pending Info', 'Settled', 'Withdrawn', 'Denied', 'Won', 'Appealed', 'In Arbitration', 'Closed'],
  GRIEVANCE_STEP: ['Informal', 'Step I', 'Step II', 'Step III', 'Mediation', 'Arbitration'],
  ISSUE_CATEGORY: ['Discipline', 'Workload', 'Scheduling', 'Pay', 'Benefits', 'Safety', 'Harassment', 'Discrimination', 'Contract Violation', 'Other'],
  ARTICLES: [
    'Art. 1 - Recognition',
    'Art. 2 - Management Rights',
    'Art. 3 - Union Rights',
    'Art. 4 - Dues Deduction',
    'Art. 5 - Non-Discrimination',
    'Art. 6 - Hours of Work',
    'Art. 7 - Overtime',
    'Art. 8 - Compensation',
    'Art. 9 - Benefits',
    'Art. 10 - Leave',
    'Art. 11 - Holidays',
    'Art. 12 - Seniority',
    'Art. 13 - Discipline',
    'Art. 14 - Safety',
    'Art. 15 - Training',
    'Art. 16 - Evaluations',
    'Art. 17 - Layoff',
    'Art. 18 - Vacancies',
    'Art. 19 - Transfers',
    'Art. 20 - Subcontracting',
    'Art. 21 - Personnel Files',
    'Art. 22 - Uniforms',
    'Art. 23 - Grievance Procedure',
    'Art. 24 - Arbitration',
    'Art. 25 - No Strike',
    'Art. 26 - Duration'
  ],
  COMM_METHODS: ['Email', 'Phone', 'Text', 'In Person']
};

/**
 * Grievance status priority order for auto-sorting
 * Lower number = higher priority (appears first in sorted list)
 * Active cases appear first, resolved cases last
 */
var GRIEVANCE_STATUS_PRIORITY = {
  'Open': 1,
  'Pending Info': 2,
  'In Arbitration': 3,
  'Appealed': 4,
  'Settled': 5,
  'Won': 6,
  'Denied': 7,
  'Withdrawn': 8,
  'Closed': 9
};

// ============================================================================
// JOB METADATA FIELDS - Maps Member Directory fields to Config dropdown sources
// ============================================================================

/**
 * Job metadata field configuration
 * Maps each Member Directory field to its corresponding Config sheet column
 * @const {Array<Object>}
 */
var JOB_METADATA_FIELDS = [
  { label: 'Job Title', memberCol: MEMBER_COLS.JOB_TITLE, configCol: CONFIG_COLS.JOB_TITLES, configName: 'Job Titles' },
  { label: 'Work Location', memberCol: MEMBER_COLS.WORK_LOCATION, configCol: CONFIG_COLS.OFFICE_LOCATIONS, configName: 'Office Locations' },
  { label: 'Unit', memberCol: MEMBER_COLS.UNIT, configCol: CONFIG_COLS.UNITS, configName: 'Units' },
  { label: 'Supervisor', memberCol: MEMBER_COLS.SUPERVISOR, configCol: CONFIG_COLS.SUPERVISORS, configName: 'Supervisors' },
  { label: 'Manager', memberCol: MEMBER_COLS.MANAGER, configCol: CONFIG_COLS.MANAGERS, configName: 'Managers' },
  { label: 'Assigned Steward', memberCol: MEMBER_COLS.ASSIGNED_STEWARD, configCol: CONFIG_COLS.STEWARDS, configName: 'Stewards' },
  { label: 'Committees', memberCol: MEMBER_COLS.COMMITTEES, configCol: CONFIG_COLS.STEWARD_COMMITTEES, configName: 'Steward Committees' },
  { label: 'Home Town', memberCol: MEMBER_COLS.HOME_TOWN, configCol: CONFIG_COLS.HOME_TOWNS, configName: 'Home Towns' }
];

/**
 * Get job metadata field config by label
 * @param {string} label - The field label (e.g., 'Job Title')
 * @returns {Object|null} Field config if found, null otherwise
 */
function getJobMetadataField(label) {
  for (var i = 0; i < JOB_METADATA_FIELDS.length; i++) {
    if (JOB_METADATA_FIELDS[i].label === label) {
      return JOB_METADATA_FIELDS[i];
    }
  }
  return null;
}

/**
 * Get job metadata field config by member column number
 * @param {number} memberCol - The member column number
 * @returns {Object|null} Field config if found, null otherwise
 */
function getJobMetadataByMemberCol(memberCol) {
  for (var i = 0; i < JOB_METADATA_FIELDS.length; i++) {
    if (JOB_METADATA_FIELDS[i].memberCol === memberCol) {
      return JOB_METADATA_FIELDS[i];
    }
  }
  return null;
}

// ============================================================================
// MULTI-SELECT COLUMN CONFIGURATION
// ============================================================================

/**
 * Columns that support multiple selections (comma-separated values)
 * Maps column number to config source column for options
 */
var MULTI_SELECT_COLS = {
  // Member Directory multi-select columns
  MEMBER_DIR: [
    { col: MEMBER_COLS.OFFICE_DAYS, configCol: CONFIG_COLS.OFFICE_DAYS, label: 'Office Days' },
    { col: MEMBER_COLS.PREFERRED_COMM, configCol: CONFIG_COLS.COMM_METHODS, label: 'Preferred Communication' },
    { col: MEMBER_COLS.BEST_TIME, configCol: CONFIG_COLS.BEST_TIMES, label: 'Best Time to Contact' },
    { col: MEMBER_COLS.COMMITTEES, configCol: CONFIG_COLS.STEWARD_COMMITTEES, label: 'Committees' },
    { col: MEMBER_COLS.ASSIGNED_STEWARD, configCol: CONFIG_COLS.STEWARDS, label: 'Assigned Steward(s)' }
  ]
};

/**
 * Check if a column in Member Directory is a multi-select column
 * @param {number} col - Column number (1-indexed)
 * @returns {Object|null} Multi-select config if found, null otherwise
 */
function getMultiSelectConfig(col) {
  for (var i = 0; i < MULTI_SELECT_COLS.MEMBER_DIR.length; i++) {
    if (MULTI_SELECT_COLS.MEMBER_DIR[i].col === col) {
      return MULTI_SELECT_COLS.MEMBER_DIR[i];
    }
  }
  return null;
}

// ============================================================================
// ID GENERATION
// ============================================================================

/**
 * Generate a name-based ID with prefix and 3 random digits
 * Format: Prefix + First 2 chars of firstName + First 2 chars of lastName + 3 random digits
 * Example: M + John Smith → MJOSM123, G + John Smith → GJOSM456
 * @param {string} prefix - ID prefix ('M' for members, 'G' for grievances)
 * @param {string} firstName - First name
 * @param {string} lastName - Last name
 * @param {Object} existingIds - Object with existing IDs as keys (for collision detection)
 * @returns {string} Generated ID (uppercase)
 */
function generateNameBasedId(prefix, firstName, lastName, existingIds) {
  var firstPart = (firstName || 'XX').substring(0, 2).toUpperCase();
  var lastPart = (lastName || 'XX').substring(0, 2).toUpperCase();
  var namePrefix = (prefix || '') + firstPart + lastPart;

  var maxAttempts = 100;
  for (var attempt = 0; attempt < maxAttempts; attempt++) {
    var randomDigits = String(Math.floor(Math.random() * 1000)).padStart(3, '0');
    var newId = namePrefix + randomDigits;

    if (!existingIds || !existingIds[newId]) {
      return newId;
    }
  }

  // Fallback: add timestamp component if too many collisions
  var timestamp = String(Date.now()).slice(-3);
  return namePrefix + timestamp;
}
