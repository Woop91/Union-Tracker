/**
 * ============================================================================
 * Constants.gs - Single Source of Truth for Dashboard Configuration
 * ============================================================================
 *
 * This file contains all configuration constants, sheet names, column mappings,
 * and deadline rules used throughout the dashboard system.
 *
 * IMPORTANT: This is the ONLY file that should contain hardcoded values.
 * All other modules should reference these constants.
 *
 * @fileoverview Central configuration for the Union Steward Dashboard
 * @version 2.0.0
 * @author Dashboard Team
 */

// ============================================================================
// SHEET NAMES - All spreadsheet tab references
// ============================================================================

/**
 * Main visible sheet names used by stewards
 */
const SHEET_NAMES = {
  MEMBER_DIRECTORY: 'Member Directory',
  GRIEVANCE_TRACKER: 'Grievance Tracker',
  CALENDAR_SYNC: 'Calendar Sync',
  REPORTS: 'Reports',
  DASHBOARD: 'Dashboard',
  SETTINGS: 'Settings',
  AUDIT_LOG: 'Audit Log'
};

/**
 * Hidden calculation sheets (prefixed with underscore)
 * These power the "self-healing" formula system
 */
const HIDDEN_SHEETS = {
  CALC_MEMBERS: '_CalcMembers',
  CALC_GRIEVANCES: '_CalcGrievances',
  CALC_DEADLINES: '_CalcDeadlines',
  CALC_STATS: '_CalcStats',
  CALC_SYNC: '_CalcSync',
  CALC_FORMULAS: '_CalcFormulas'
};

// ============================================================================
// COLUMN MAPPINGS - Zero-indexed column positions
// ============================================================================

/**
 * Member Directory column structure
 */
const MEMBER_COLUMNS = {
  ID: 0,
  FIRST_NAME: 1,
  LAST_NAME: 2,
  EMPLOYEE_ID: 3,
  DEPARTMENT: 4,
  JOB_TITLE: 5,
  HIRE_DATE: 6,
  SENIORITY_DATE: 7,
  EMAIL: 8,
  PHONE: 9,
  STATUS: 10,
  UNION_STATUS: 11,
  NOTES: 12,
  LAST_UPDATED: 13
};

/**
 * Grievance Tracker column structure
 */
const GRIEVANCE_COLUMNS = {
  GRIEVANCE_ID: 0,
  MEMBER_ID: 1,
  MEMBER_NAME: 2,
  FILING_DATE: 3,
  GRIEVANCE_TYPE: 4,
  ARTICLE_VIOLATED: 5,
  DESCRIPTION: 6,
  CURRENT_STEP: 7,
  STEP_1_DATE: 8,
  STEP_1_DUE: 9,
  STEP_1_STATUS: 10,
  STEP_2_DATE: 11,
  STEP_2_DUE: 12,
  STEP_2_STATUS: 13,
  STEP_3_DATE: 14,
  STEP_3_DUE: 15,
  STEP_3_STATUS: 16,
  ARBITRATION_DATE: 17,
  RESOLUTION: 18,
  OUTCOME: 19,
  DRIVE_FOLDER: 20,
  NOTES: 21,
  STATUS: 22,
  LAST_UPDATED: 23
};

// ============================================================================
// GRIEVANCE DEADLINE RULES - Based on Article 23A
// ============================================================================

/**
 * Deadline configuration for grievance steps (in calendar days)
 * These values are derived from the collective bargaining agreement
 */
const DEADLINE_RULES = {
  // Step 1: Informal resolution attempt
  STEP_1: {
    DAYS_TO_FILE: 14,        // Days from incident to file Step 1
    DAYS_FOR_RESPONSE: 7,    // Days for management response
    DESCRIPTION: 'Informal Discussion with Supervisor'
  },

  // Step 2: Formal written grievance
  STEP_2: {
    DAYS_TO_APPEAL: 7,       // Days to appeal Step 1 denial
    DAYS_FOR_RESPONSE: 14,   // Days for management response
    DESCRIPTION: 'Written Grievance to Department Head'
  },

  // Step 3: Third-level review
  STEP_3: {
    DAYS_TO_APPEAL: 10,      // Days to appeal Step 2 denial
    DAYS_FOR_RESPONSE: 21,   // Days for management response
    DESCRIPTION: 'Review by Labor Relations'
  },

  // Arbitration
  ARBITRATION: {
    DAYS_TO_DEMAND: 30,      // Days to demand arbitration after Step 3
    DESCRIPTION: 'Binding Arbitration'
  }
};

/**
 * Grievance status values
 */
const GRIEVANCE_STATUS = {
  OPEN: 'Open',
  PENDING: 'Pending Response',
  APPEALED: 'Appealed',
  RESOLVED: 'Resolved',
  DENIED: 'Denied',
  WITHDRAWN: 'Withdrawn',
  AT_ARBITRATION: 'At Arbitration',
  CLOSED: 'Closed'
};

/**
 * Grievance outcome values
 */
const GRIEVANCE_OUTCOMES = {
  SUSTAINED: 'Sustained',
  DENIED: 'Denied',
  SETTLED: 'Settled',
  WITHDRAWN: 'Withdrawn',
  PENDING: 'Pending'
};

// ============================================================================
// UI CONFIGURATION
// ============================================================================

/**
 * Theme colors for the dashboard UI
 */
const UI_THEME = {
  PRIMARY_COLOR: '#1a73e8',
  SECONDARY_COLOR: '#34a853',
  WARNING_COLOR: '#fbbc04',
  DANGER_COLOR: '#ea4335',
  BACKGROUND: '#ffffff',
  TEXT_PRIMARY: '#202124',
  TEXT_SECONDARY: '#5f6368',
  BORDER_COLOR: '#dadce0'
};

/**
 * Dialog dimensions
 */
const DIALOG_SIZES = {
  SMALL: { width: 400, height: 300 },
  MEDIUM: { width: 600, height: 450 },
  LARGE: { width: 800, height: 600 },
  FULLSCREEN: { width: 1000, height: 700 }
};

/**
 * Sidebar configuration
 */
const SIDEBAR_CONFIG = {
  TITLE: 'Union Dashboard',
  WIDTH: 300
};

// ============================================================================
// INTEGRATION SETTINGS
// ============================================================================

/**
 * Google Drive folder structure
 */
const DRIVE_CONFIG = {
  ROOT_FOLDER_NAME: 'Union Grievance Files',
  SUBFOLDER_TEMPLATE: 'Grievance_{grievanceId}_{memberName}',
  ALLOWED_MIME_TYPES: [
    'application/pdf',
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'image/jpeg',
    'image/png'
  ],
  MAX_FILE_SIZE_MB: 25
};

/**
 * Google Calendar configuration
 */
const CALENDAR_CONFIG = {
  CALENDAR_NAME: 'Grievance Deadlines',
  EVENT_COLOR: '11',           // Red color ID
  REMINDER_DAYS: [7, 3, 1],    // Days before deadline to remind
  DEFAULT_DURATION_HOURS: 1
};

// ============================================================================
// SYSTEM SETTINGS
// ============================================================================

/**
 * Batch processing limits to avoid timeouts
 */
const BATCH_LIMITS = {
  MAX_ROWS_PER_BATCH: 100,
  MAX_API_CALLS_PER_BATCH: 50,
  PAUSE_BETWEEN_BATCHES_MS: 500,
  MAX_EXECUTION_TIME_MS: 300000  // 5 minutes
};

/**
 * Data validation rules
 */
const VALIDATION_RULES = {
  PHONE_PATTERN: /^\(?[\d]{3}\)?[-.\s]?[\d]{3}[-.\s]?[\d]{4}$/,
  EMAIL_PATTERN: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
  EMPLOYEE_ID_PATTERN: /^[A-Z]{2}\d{6}$/,
  GRIEVANCE_ID_PATTERN: /^GRV-\d{4}-\d{4}$/
};

/**
 * Audit log event types
 */
const AUDIT_EVENTS = {
  GRIEVANCE_CREATED: 'GRIEVANCE_CREATED',
  GRIEVANCE_UPDATED: 'GRIEVANCE_UPDATED',
  GRIEVANCE_STEP_ADVANCED: 'GRIEVANCE_STEP_ADVANCED',
  MEMBER_ADDED: 'MEMBER_ADDED',
  MEMBER_UPDATED: 'MEMBER_UPDATED',
  CALENDAR_SYNCED: 'CALENDAR_SYNCED',
  FOLDER_CREATED: 'FOLDER_CREATED',
  SYSTEM_REPAIR: 'SYSTEM_REPAIR',
  SETTINGS_CHANGED: 'SETTINGS_CHANGED'
};

/**
 * User permission levels
 */
const PERMISSION_LEVELS = {
  VIEWER: 0,
  STEWARD: 1,
  CHIEF_STEWARD: 2,
  ADMIN: 3
};

// ============================================================================
// HELPER FUNCTIONS FOR CONSTANTS
// ============================================================================

/**
 * Gets the full list of all sheet names (visible and hidden)
 * @return {string[]} Array of all sheet names
 */
function getAllSheetNames() {
  return [...Object.values(SHEET_NAMES), ...Object.values(HIDDEN_SHEETS)];
}

/**
 * Validates a grievance ID format
 * @param {string} id - The ID to validate
 * @return {boolean} True if valid format
 */
function isValidGrievanceId(id) {
  return VALIDATION_RULES.GRIEVANCE_ID_PATTERN.test(id);
}

/**
 * Generates a new grievance ID based on current year and sequence
 * @param {number} sequence - The sequence number for this year
 * @return {string} Formatted grievance ID
 */
function generateGrievanceId(sequence) {
  const year = new Date().getFullYear();
  const paddedSequence = String(sequence).padStart(4, '0');
  return `GRV-${year}-${paddedSequence}`;
}

/**
 * Gets deadline days for a specific grievance step
 * @param {number} step - The step number (1-4, where 4 is arbitration)
 * @return {Object} Deadline configuration for the step
 */
function getDeadlineConfig(step) {
  const configs = {
    1: DEADLINE_RULES.STEP_1,
    2: DEADLINE_RULES.STEP_2,
    3: DEADLINE_RULES.STEP_3,
    4: DEADLINE_RULES.ARBITRATION
  };
  return configs[step] || null;
}
