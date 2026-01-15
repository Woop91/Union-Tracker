# 509 Dashboard - Architecture & Implementation Reference

**Version:** 3.6.0 / v3.48 (Strategic Command Center, Midnight Auto-Refresh, Dynamic Config)
**Last Updated:** 2026-01-15
**Purpose:** Union grievance tracking and member engagement system for SEIU Local 509

---

## Creator & License

**Creator & Owner:** Wardis N. Vizcaino
**Role:** Steward at SEIU Local 509
**Contact:** wardis@pm.me

**License:** Free for use by non-profit collective bargaining groups and unions. No license required.

---

## Quick Start

> ⚠️ **IMPORTANT: Deploy ONLY `ConsolidatedDashboard.gs`**
> The modular `.gs` files are source files used to generate ConsolidatedDashboard.gs.
> Deploying multiple files will cause function conflicts and trigger errors.

1. Copy **only** `ConsolidatedDashboard.gs` to Google Apps Script
2. Run `CREATE_509_DASHBOARD()` to create 5 sheets + 6 hidden calculation sheets
3. Use `🔧 Admin > 🎭 Demo Data > Seed All Sample Data` to populate test data
4. Customize Config sheet with your organization's values

---

## ⚠️ Protected Code - DO NOT MODIFY

The following code sections are **USER APPROVED** and should **NOT be modified or removed**:

### Dashboard Modal Popup (formerly Custom View)

**Location:** `ConsolidatedDashboard.gs` and `MobileQuickActions.gs`

**Protected Functions:**
| Function | Purpose |
|----------|---------|
| `showInteractiveDashboardTab()` | Opens the modal dialog popup (titled "📊 Dashboard") |
| `getInteractiveDashboardHtml()` | Returns the HTML/CSS/JS for the tabbed UI |
| `getInteractiveOverviewData()` | Fetches overview statistics |
| `getInteractiveGrievanceData()` | Fetches grievance list data |
| `getInteractiveMemberData()` | Fetches member list data |
| `getMyStewardCases()` | Fetches steward's assigned grievances for My Cases tab |

**Features:**
- **My Cases Tab** (NEW) - Stewards can view their assigned grievances with stats and filtering
- Overview, Grievances, Members, Analytics tabs

**Why Protected:** This modal provides essential dashboard functionality that users rely on for quick access to data.

### Member Satisfaction Dashboard Modal (Enhanced v2.0)

**Location:** `ConsolidatedDashboard.gs`

**Functions:**
| Function | Purpose |
|----------|---------|
| `showSatisfactionDashboard()` | Opens the 900x750 modal dialog |
| `getSatisfactionDashboardHtml()` | Returns HTML with 5-tab interface |
| `getSatisfactionOverviewData()` | Fetches overview stats, Member Advocacy Index, insights |
| `getSatisfactionResponseData()` | Fetches individual survey responses |
| `getSatisfactionSectionData()` | Fetches scores by survey section |
| `getSatisfactionAnalyticsData()` | Fetches worksite/role analysis, priorities |
| `getSatisfactionTrendData()` | NEW: Fetches trend data by time period |
| `getSatisfactionLocationDrill()` | NEW: Fetches drill-down data for specific worksites |

**Tabs:**
1. **Overview** - Key metrics (total responses, avg satisfaction, Member Advocacy Index, response rate), gauge charts, auto-generated insights
2. **Trends** (NEW) - Line chart for satisfaction trends over time, responses by month, time period filters (All Time/Year/90 Days/30 Days), donut chart for top issues
3. **Responses** - Searchable list with satisfaction level filtering (High/Medium/Needs Attention)
4. **Sections** - Bar chart of all 11 survey sections ranked by score
5. **Insights** - Clickable worksite breakdowns with drill-down modals, steward contact impact analysis, top member priorities

**Enhancements (v2.0):**
- **Renamed "Loyalty Score" to "Member Advocacy Index"** for clearer understanding
- **Added Trends tab** with line charts and time period filtering
- **Added donut charts** for issue distribution visualization
- **Added clickable worksite drill-down** - click any worksite bar to see detailed responses
- **Enhanced color palette** - teal, indigo, pink, amber accent colors
- **Clickable stats** - response count navigates to Responses tab

**Why Added:** Provides interactive survey analysis without leaving the spreadsheet, matching the Dashboard modal pattern.

---

## No Formulas in Visible Sheets Architecture

**Design Principle:** All visible sheets (Dashboard, Member Satisfaction, Feedback) contain only VALUES, never formulas. This provides:

1. **Self-Healing Data** - Values are recomputed by JavaScript on each data change
2. **Reliability** - No broken formula references when rows are added/deleted
3. **Performance** - No circular dependency calculations or formula chains
4. **Consistency** - All data flows through controlled sync functions

### How It Works

| Visible Sheet | Sync Function | When Called |
|---------------|---------------|-------------|
| Dashboard | `syncDashboardValues()` | On Grievance Log or Member Directory edit |
| Member Satisfaction | `syncSatisfactionValues()` | On form submission, sheet creation |
| Feedback | `syncFeedbackValues()` | On Feedback sheet edit |

### Data Flow

```
User Edit → onEditAutoSync() → syncDashboardValues() → writeDashboardValues_()
                                                      └── computeDashboardMetrics_()
Form Submit → onSatisfactionFormSubmit() → computeSatisfactionRowAverages()
                                         → syncSatisfactionValues()
```

### Member Directory: Start Grievance Checkbox

When the "Start Grievance" checkbox (column AE) is checked in Member Directory:
1. `onEditAutoSync()` detects the checkbox change
2. Checkbox is immediately unchecked (so it can be reused)
3. `openGrievanceFormForRow_()` opens a pre-filled grievance form with member data
4. User completes form → `onGrievanceFormSubmit()` adds to Grievance Log

---

## File Architecture

### Modular Architecture (9 Source Files)

This repository implements a **streamlined 9-file modular architecture** following the Separation of Concerns principle. Each module handles a specific aspect of the dashboard functionality.

```
MULTIPLE-SCRIPS-REPO/
├── src/                        # Source files for build (9 modules)
│   ├── 01_Constants.gs         # Configuration constants (SHEETS, COLORS, MEMBER_COLS, GRIEVANCE_COLS, COMMAND_CONFIG)
│   ├── 02_MemberManager.gs     # Member operations, steward management, ID generation
│   ├── 03_GrievanceManager.gs  # Grievance lifecycle, deadlines, step advancement
│   ├── 04_UIService.gs         # UI, Comfort View, mobile, Strategic Command Center dashboards
│   ├── 05_Integrations.gs      # Google Drive, Calendar, WebApp, email notifications
│   ├── 06_Maintenance.gs       # Admin tools, diagnostics, caching, validation
│   ├── 07_DevTools.gs          # Demo data seeding - DELETE BEFORE PRODUCTION
│   ├── 08_Code.gs              # Core setup, hidden sheets, dashboard creation, multi-select
│   ├── 09_Main.gs              # Entry point, onOpen, onEdit triggers
│   └── MultiSelectDialog.html  # HTML template for multi-select UI
├── dist/                       # Build output
│   └── ConsolidatedDashboard.gs # Combined file for deployment (auto-generated)
├── AIR.md                      # Architecture & Implementation Reference (this document)
├── README.md                   # Quick start guide
├── QUICK_DEPLOY.md             # 5-minute deployment guide
├── build.js                    # Node.js build script
├── package.json                # npm configuration
└── appsscript.json             # Google Apps Script manifest
```

### Build Process

The modular files are concatenated into a single `ConsolidatedDashboard.gs` for deployment:

```bash
npm install          # Install dependencies (first time only)
npm run build        # Build ConsolidatedDashboard.gs
npm run watch        # Watch mode for development
```

### File Descriptions (9-File Architecture)

**01_Constants.gs** (~700 lines) - Configuration & Column Mapping
- `SHEETS` - Sheet name constants (3 data + 2 dashboard + 6 hidden)
- `COLORS` - Brand color scheme
- `MEMBER_COLS` - 31 Member Directory column positions
- `GRIEVANCE_COLS` - 34 Grievance Log column positions
- `CONFIG_COLS` - Config sheet column positions (includes Strategic Command settings at AS-AW)
- `COMMAND_CONFIG` - Strategic Command Center configuration (v3.6.0):
  - `SYSTEM_NAME`, `VERSION` - System identification
  - `ESCALATION_STATUSES`, `ESCALATION_STEPS` - Alert triggers
  - `UNIT_CODES` - Dynamic from Config sheet
  - `THEME` - Roboto theme settings (header colors, alt rows, fonts)
  - `STATUS_COLORS` - Status-based coloring (Open=Yellow, Won=Green, Denied=Red, etc.)
- `DEFAULT_CONFIG` - Default dropdown values
- `MULTI_SELECT_COLS` - Configuration for multi-select columns
- `getMultiSelectConfig()` - Get multi-select config for a column
- `generateNameBasedId(prefix, firstName, lastName, existingIds)` - Generate unique ID
  - **Format:** `PREFIX + first 2 chars of firstName + first 2 chars of lastName + 3 random digits`
  - **Member ID:** `MJOSM123` (M + "John Smith" → JO + SM + 123)
  - **Grievance ID:** `GJOSM456` (G + "John Smith" → JO + SM + 456)
  - Includes collision detection to ensure uniqueness
- `getColumnLetter()` - Convert column number to letter
- `getColumnNumber()` - Convert column letter to number
- `mapMemberRow()` - Map row array to member object
- `mapGrievanceRow()` - Map row array to grievance object
- `getMemberHeaders()` - Get all 31 member column headers
- `getGrievanceHeaders()` - Get all 34 grievance column headers

**02_MemberManager.gs** (~350 lines) - Member Directory Operations
- Member CRUD Operations:
  - `addMember()` - Add new member to directory
  - `updateMember()` - Update existing member data
  - `getMemberById()` - Lookup member by ID
  - `searchMembers()` - Search members by name/email/ID
- Member ID Generation:
  - `generateMemberID_()` - Generate unique Member ID (M + name prefix + digits)
- Steward Management:
  - `getAllStewards()` - Get all active stewards
  - `getStewardWorkload()` - Calculate steward case loads with win rates
- Data Sync:
  - `syncMemberGrievanceData()` - Sync grievance counts to Member Directory

**03_GrievanceManager.gs** (~1048 lines) - Grievance Lifecycle Management
- `startNewGrievance()` - Open pre-filled grievance form
- `advanceGrievanceStep()` - Advance grievance to next step
- `recalcAllGrievancesBatched()` - Batch recalculate grievance deadlines
- `onGrievanceFormSubmit()` - Handle form submission trigger
- `calculateInitialDeadlines()` - Calculate filing and response deadlines
- `getGrievanceById()` - Lookup grievance by ID
- `bulkUpdateGrievanceStatus()` - Bulk update grievance statuses
- `resolveGrievance()` - Close grievance with outcome

**04_UIService.gs** (~4500 lines) - UI, Mobile, Comfort View & Strategic Command Center
*Consolidated from: UIService, ComfortViewFeatures, MobileQuickActions, StrategicCommandCenter*

- UI Components & Dialogs:
  - `showDesktopSearch()` - Desktop search dialog
  - `showMultiSelectDialog()` - Multi-select checkbox dialog
  - `showToast()`, `showAlert()` - Notification helpers
  - `getCommonStyles()` - Shared CSS for dialogs
  - `showDashboardSidebar()` - Sidebar panel for extended UI
  - `showAdvancedSearch()` - Advanced search with filters

- Comfort View Accessibility:
  - `showComfortViewPanel()` - Main Comfort View settings panel
  - `getComfortViewSettings()`, `saveComfortViewSettings()`, `resetComfortViewSettings()` - Settings management
  - `applyComfortViewSettings()` - Apply visual settings
  - `activateFocusMode()`, `deactivateFocusMode()` - Focus mode (hide non-essential sheets)
  - `toggleZebraStripes()`, `applyZebraStripes()`, `removeZebraStripes()` - Row banding
  - `toggleGridlines()`, `hideAllGridlines()`, `showAllGridlines()` - Gridline control
  - `showQuickCaptureNotepad()` - Quick note-taking dialog
  - `startPomodoroTimer()` - Built-in pomodoro timer
  - `setBreakReminders()`, `showBreakReminder()` - Break notification system
  - `showThemeManager()` - Theme selection UI
  - `applyTheme()`, `applyThemeToSheet()`, `previewTheme()` - Theme application

- Mobile Interface:
  - `showMobileDashboard()` - Touch-optimized dashboard
  - `getMobileDashboardStats()` - Dashboard statistics
  - `getRecentGrievancesForMobile()` - Recent grievances (limit)
  - `showMobileGrievanceList()` - Mobile grievance list
  - `showMobileUnifiedSearch()` - Mobile search UI
  - `getMobileSearchData()` - Search handler
  - `showMyAssignedGrievances()` - View user's assigned cases

- Quick Actions:
  - `showQuickActionsMenu()` - Context-aware quick actions
  - `showMemberQuickActions()` - Quick actions for member row
  - `showGrievanceQuickActions()` - Quick actions for grievance row
  - `quickUpdateGrievanceStatus()` - One-click status update
  - `composeEmailForMember()` - Email composition dialog
  - `sendQuickEmail()` - Send email via MailApp
  - `emailSurveyToMember()` - Email satisfaction survey link
  - `emailContactFormToMember()` - Email contact update form link
  - `emailDashboardLinkToMember()` - Email dashboard access link
  - `emailGrievanceStatusToMember()` - Email grievance status update

- Strategic Command Center (v3.6.0):
  - `rebuildExecutiveDashboard()` - Executive Command (PII) dashboard
  - `rebuildMemberAnalytics()` - Member Analytics (No PII) dashboard
  - `generateUnitHotZones()` - Identify locations with 3+ active grievances
  - `identifyRisingStars()` - Top-performing stewards by score
  - `generateManagementHostilityReport()` - Analyze denial rates
  - `generateBargainingCheatSheet()` - Strategic data for negotiations
  - `generateMissingMemberIDs()` - Auto-ID generator with unit codes
  - `checkDuplicateIDs()` - Find duplicate Member IDs
  - `createGrievancePDF()` - PDF with digital signature block
  - `promoteToSteward()`, `demoteFromSteward()` - Steward management
  - `applyGlobalStyling()` - Roboto theme with zebra stripes
  - `setupMidnightTrigger()`, `removeMidnightTrigger()` - Daily refresh automation
  - `midnightAutoRefresh()` - 12AM dashboard refresh + overdue alerts

**05_Integrations.gs** (~1200 lines) - External Services & WebApp
*Consolidated from: Integrations, WebApp*

- Google Drive Integration:
  - `setupDriveFolderForGrievance()` - Create grievance folder with subfolders
  - `batchCreateGrievanceFolders()` - Bulk folder creation
  - `getOrCreateDeadlinesCalendar()` - Get/create calendar
- Calendar Integration:
  - `syncDeadlinesToCalendar()` - Sync all grievance deadlines
  - `clearAllCalendarEvents()` - Clear dashboard events
- Email Notifications:
  - `sendDeadlineReminders()` - Send upcoming deadline alerts
  - `sendEmailToMember()` - Send email to specific member
- Web App Entry Point:
  - `doGet(e)` - Web app entry point, routes to pages
  - `getWebAppDashboardHtml()` - Main dashboard with stats
  - `getWebAppSearchHtml()` - Full-text search page
  - `getWebAppGrievanceListHtml()` - Filterable grievance list
  - `getWebAppMemberListHtml()` - Member list page
  - `getWebAppLinksHtml()` - Forms and resources links page
  - `showWebAppUrl()` - Display web app URL after deployment

**06_Maintenance.gs** (~2500 lines) - Admin Tools, Diagnostics, Caching & Validation
*Consolidated from: Maintenance, PerformanceUndo, DataIntegrity, TestingValidation*

- Diagnostics & Repair:
  - `DIAGNOSE_SETUP()` - System health check
  - `REPAIR_DASHBOARD()` - Repair hidden sheets and triggers
  - `logAuditEvent()` - Log action to audit sheet
  - `showDiagnosticsDialog()` - Display diagnostics UI
  - `NUCLEAR_RESET_HIDDEN_SHEETS()` - Complete hidden sheet reset

- Caching Layer:
  - `getCachedData()` - Get data from cache or load
  - `setCachedData()` - Store data in cache
  - `invalidateCache()`, `invalidateAllCaches()` - Clear caches
  - `warmUpCaches()` - Pre-populate caches
  - `getCachedGrievances()`, `getCachedMembers()`, `getCachedStewards()` - Cached getters
  - `showCacheStatusDashboard()` - Cache status UI

- Undo/Redo System:
  - `getUndoHistory()`, `saveUndoHistory()` - History management
  - `recordAction()`, `recordCellEdit()`, `recordRowAddition()` - Action recording
  - `undoLastAction()`, `redoLastAction()` - Undo/redo operations
  - `showUndoRedoPanel()` - Undo/redo UI
  - `createGrievanceSnapshot()`, `restoreFromSnapshot()` - Full snapshot backup

- Batch Operations:
  - `batchSetValues()` - Write multiple values in single operation
  - `batchSetRowValues()` - Update multiple columns efficiently
  - `batchAppendRows()` - Add multiple rows at once
  - `executeWithRetry()` - Retry with exponential backoff

- Data Integrity:
  - `checkDuplicateMemberId()` - Duplicate Member ID detection
  - `findOrphanedGrievances()` - Find grievances with invalid Member IDs
  - `highlightOrphanedGrievances()` - Visual highlighting of issues
  - `calculateStewardWorkload()` - Load scores per steward
  - `showStewardWorkloadDashboard()` - Workload dashboard
  - `findMissingConfigValues()` - Scan for missing Config values
  - `autoFixMissingConfigValues()` - Auto-add missing values
  - `archiveClosedGrievances()` - Auto-archive old cases

- Testing Framework:
  - `Assert` - Assertion library (assertEquals, assertTrue, etc.)
  - `runAllTests()`, `runQuickTests()` - Test suite runners
  - `generateTestReport()` - Create test results sheet
  - `VALIDATION_PATTERNS` - Regex patterns for Member ID, Grievance ID, Email, Phone
  - `validateEmailAddress()`, `validatePhoneNumber()` - Format validation
  - `runBulkValidation()` - Validate all data
  - `showValidationReport()` - Display validation issues

**07_DevTools.gs** (~1400 lines) - Demo Data Generation (DELETE BEFORE PRODUCTION)
- `SEED_SAMPLE_DATA()` - Seeds Config + 1,000 members + 300 grievances + 50 survey responses
- `seedConfigData()` - Populate Config dropdowns
- `seedSatisfactionData()` - Seed sample survey responses
- `seedFeedbackData()` - Seed sample feedback entries
- `SEED_MEMBERS(count, grievancePercent)` - Seed N members with grievances
- `SEED_GRIEVANCES(count)` - Seed N grievances for existing members
- `SEED_MEMBERS_DIALOG()`, `SEED_MEMBERS_ADVANCED_DIALOG()` - Prompts for counts
- `generateSingleMemberRow()` - Generate one member row (31 columns)
- `generateSingleGrievanceRow()` - Generate one grievance row (34 columns)
- `NUKE_SEEDED_DATA()` - Clear seeded data with confirmation
- `NUKE_CONFIG_DROPDOWNS()` - Clear only Config dropdowns
- Helper functions: `getConfigValues()`, `randomChoice()`, `randomDate()`, `addDays()`

**08_Code.gs** (~5000 lines) - Core Setup, Hidden Sheets & Dashboard Creation
*Consolidated from: Code, HiddenSheets, FormulaService*

- Main Setup:
  - `CREATE_509_DASHBOARD()` - Main setup function (creates sheets + hidden)
  - `setupDataValidations()` - Apply dropdown validations
  - `setDropdownValidation()`, `setMultiSelectValidation()` - Validation helpers
  - `showMultiSelectDialog()`, `applyMultiSelectValue()` - Multi-select UI
  - `getOrCreateSheet()` - Helper: get or create sheet
  - `rebuildDashboard()` - Refresh data and validations
  - `refreshAllFormulas()` - Refresh all formulas and sync

- Sheet Creation (11 functions):
  - `createConfigSheet()`, `createMemberDirectory()`, `createGrievanceLog()`
  - `createDashboard()`, `createInteractiveDashboard()`
  - `createSatisfactionSheet()`, `createFeedbackSheet()`
  - `createFunctionChecklistSheet_()` - Includes Phase 14: Strategic Command Center
  - `createGettingStartedSheet()`, `createFAQSheet()`, `createConfigGuideSheet()`

- Hidden Sheet Management:
  - `setupAllHiddenSheets()` - Create all 6 hidden calculation sheets
  - `setupGrievanceCalcSheet()` - Grievance timeline formulas
  - `setupGrievanceFormulasSheet()` - Member lookup formulas
  - `setupMemberLookupSheet()` - Member → Grievance sync
  - `setupStewardContactCalcSheet()` - Steward contact tracking
  - `setupDashboardCalcSheet()` - Dashboard summary metrics
  - `setupStewardPerformanceCalcSheet()` - Performance scores

- Value Sync Functions (No Formulas in Visible Sheets):
  - `syncDashboardValues()` - Compute and write Dashboard metrics as VALUES
  - `computeDashboardMetrics_()` - Calculate 100+ Dashboard metrics
  - `syncSatisfactionValues()` - Compute Member Satisfaction metrics
  - `syncFeedbackValues()` - Compute Feedback sheet metrics
  - `syncAllData()` - Sync all cross-sheet data
  - `syncGrievanceToMemberDirectory()` - Sync grievance data to members (AB-AD)
  - `syncMemberToGrievanceLog()` - Sync member data to grievances
  - `sortGrievanceLogByStatus()` - Auto-sort by status priority

- Trigger & Repair:
  - `onEditAutoSync()` - Auto-sync trigger handler
  - `installAutoSyncTrigger()`, `removeAutoSyncTrigger()` - Trigger management
  - `repairAllHiddenSheets()` - Self-healing repair function
  - `verifyHiddenSheets()` - Verification and diagnostics

- Form Workflows:
  - Grievance Form: `startNewGrievance()`, `onGrievanceFormSubmit()`, `buildGrievanceFormUrl_()`
  - Contact Form: `sendContactInfoForm()`, `onContactFormSubmit()`
  - Satisfaction Survey: `getSatisfactionSurveyLink()`, `onSatisfactionFormSubmit()`
  - `saveFormUrlsToConfig()`, `getFormUrlFromConfig()` - Form URL management

- Google Drive/Calendar/Email Integration:
  - `setupDriveFolderForGrievance()`, `batchCreateGrievanceFolders()`
  - `syncDeadlinesToCalendar()`, `showUpcomingDeadlinesFromCalendar()`
  - `showNotificationSettings()`, `testDeadlineNotifications()`

- Audit Log:
  - `setupAuditLogSheet()`, `logAuditEvent()`, `onEditAudit()`, `viewAuditLog()`

- Member Satisfaction Dashboard:
  - `showSatisfactionDashboard()` - 900x750 modal with 5-tab interface
  - `getSatisfactionOverviewData()`, `getSatisfactionResponseData()`
  - `getSatisfactionSectionData()`, `getSatisfactionAnalyticsData()`
  - `getSatisfactionTrendData()`, `getSatisfactionLocationDrill()`

**09_Main.gs** (~900 lines) - Entry Point & Triggers
- Menu System:
  - `onOpen()` - Create 6 menus (509 Dashboard, Grievances, View, Settings, Admin, 509 Command)
- Edit Handlers:
  - `onEdit()` - Handle cell edit events with stage-gate workflow
  - `handleStageGateWorkflow_()` - Automatic escalation alerts on step changes
  - `sendEscalationAlert_()` - Email alerts to Chief Steward
- Config Helpers:
  - `getConfigValue_()` - Read values from Config sheet columns
  - `getEscalationStatuses_()`, `getEscalationSteps_()` - Dynamic escalation config
  - `getUnitCodes_()` - Parse unit codes from Config
- Initialization:
  - `initializeDashboard()` - First-time initialization
  - `setupTriggers()` - Setup all script triggers
  - `dailyTrigger()` - Daily automated tasks
- Member Actions:
  - `startGrievanceForMember()` - Start grievance from Member Directory

---

## Column Mapping System

**CRITICAL: ALL column references must use these constants, never hardcoded letters.**

### MEMBER_COLS (31 columns: A-AE)

```javascript
var MEMBER_COLS = {
  // Section 1: Identity & Core Info (A-D)
  MEMBER_ID: 1,           // A
  FIRST_NAME: 2,          // B
  LAST_NAME: 3,           // C
  JOB_TITLE: 4,           // D

  // Section 2: Location & Work (E-G)
  WORK_LOCATION: 5,       // E
  UNIT: 6,                // F
  OFFICE_DAYS: 7,         // G - Multi-select

  // Section 3: Contact Information (H-K)
  EMAIL: 8,               // H
  PHONE: 9,               // I
  PREFERRED_COMM: 10,     // J - Multi-select
  BEST_TIME: 11,          // K - Multi-select

  // Section 4: Organizational Structure (L-P)
  SUPERVISOR: 12,         // L
  MANAGER: 13,            // M
  IS_STEWARD: 14,         // N
  COMMITTEES: 15,         // O - Multi-select
  ASSIGNED_STEWARD: 16,   // P - Multi-select

  // Section 5: Engagement Metrics (Q-T)
  LAST_VIRTUAL_MTG: 17,   // Q
  LAST_INPERSON_MTG: 18,  // R
  OPEN_RATE: 19,          // S
  VOLUNTEER_HOURS: 20,    // T

  // Section 6: Member Interests (U-X)
  INTEREST_LOCAL: 21,     // U
  INTEREST_CHAPTER: 22,   // V
  INTEREST_ALLIED: 23,    // W
  HOME_TOWN: 24,          // X

  // Section 7: Steward Contact Tracking (Y-AA)
  RECENT_CONTACT_DATE: 25, // Y
  CONTACT_STEWARD: 26,    // Z
  CONTACT_NOTES: 27,      // AA

  // Section 8: Grievance Management (AB-AE)
  HAS_OPEN_GRIEVANCE: 28, // AB - auto-sync: "Yes"/"No" from Grievance Log
  GRIEVANCE_STATUS: 29,   // AC - auto-sync: Status from Grievance Log
  NEXT_DEADLINE: 30,      // AD - auto-sync: Days to Deadline (number or "Overdue")
  START_GRIEVANCE: 31     // AE (checkbox)
};
```

### GRIEVANCE_COLS (34 columns: A-AH)

```javascript
var GRIEVANCE_COLS = {
  // Section 1: Identity (A-D)
  GRIEVANCE_ID: 1,        // A
  MEMBER_ID: 2,           // B
  FIRST_NAME: 3,          // C
  LAST_NAME: 4,           // D

  // Section 2: Status & Assignment (E-F)
  STATUS: 5,              // E
  CURRENT_STEP: 6,        // F

  // Section 3: Timeline - Filing (G-I)
  INCIDENT_DATE: 7,       // G
  FILING_DEADLINE: 8,     // H (auto-calc: +21 days)
  DATE_FILED: 9,          // I

  // Section 4: Timeline - Step I (J-K)
  STEP1_DUE: 10,          // J (auto-calc: +30 days)
  STEP1_RCVD: 11,         // K

  // Section 5: Timeline - Step II (L-O)
  STEP2_APPEAL_DUE: 12,   // L (auto-calc: +10 days)
  STEP2_APPEAL_FILED: 13, // M
  STEP2_DUE: 14,          // N (auto-calc: +30 days)
  STEP2_RCVD: 15,         // O

  // Section 6: Timeline - Step III (P-R)
  STEP3_APPEAL_DUE: 16,   // P (auto-calc: +30 days)
  STEP3_APPEAL_FILED: 17, // Q
  DATE_CLOSED: 18,        // R

  // Section 7: Calculated Metrics (S-U)
  DAYS_OPEN: 19,          // S (auto-calc)
  NEXT_ACTION_DUE: 20,    // T (auto-calc)
  DAYS_TO_DEADLINE: 21,   // U (auto-calc)

  // Section 8: Case Details (V-W)
  ARTICLES: 22,           // V
  ISSUE_CATEGORY: 23,     // W

  // Section 9: Contact & Location (X-AA)
  MEMBER_EMAIL: 24,       // X
  UNIT: 25,               // Y
  LOCATION: 26,           // Z
  STEWARD: 27,            // AA

  // Section 10: Resolution (AB)
  RESOLUTION: 28,         // AB

  // Section 11: Coordinator Notifications (AC-AF)
  MESSAGE_ALERT: 29,      // AC (checkbox)
  COORDINATOR_MESSAGE: 30, // AD
  ACKNOWLEDGED_BY: 31,    // AE
  ACKNOWLEDGED_DATE: 32,  // AF

  // Section 12: Drive Integration (AG-AH)
  DRIVE_FOLDER_ID: 33,    // AG
  DRIVE_FOLDER_URL: 34    // AH
};
```

### SATISFACTION_COLS (68-Question Google Form Response + Summary)

**Form Response Area (A-BP):** Auto-populated by linked Google Form (68 questions + timestamp)
**Summary/Chart Data Area (BT-CD):** Section averages for charts and dashboards

```javascript
var SATISFACTION_COLS = {
  // Form Response Columns (auto-created by Google Form)
  TIMESTAMP: 1,                   // A - Auto by Google Forms

  // Work Context (Q1-5)
  Q1_WORKSITE: 2,                 // B - Worksite/Program/Region
  Q2_ROLE: 3,                     // C - Role/Job Group
  Q3_SHIFT: 4,                    // D - Day/Evening/Night/Rotating
  Q4_TIME_IN_ROLE: 5,             // E - Tenure range
  Q5_STEWARD_CONTACT: 6,          // F - Yes/No (branching)

  // Overall Satisfaction Q6-9, Steward Ratings Q10-17, etc.
  // ... (68 total questions, see Constants.gs for full list)

  // Summary/Chart Data Area (BT onwards)
  SUMMARY_START: 72,              // BT
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
  AVG_SCHEDULING: 82              // CD - Avg of Q56-Q62
};

// Survey Section Definitions for Analysis
var SATISFACTION_SECTIONS = {
  WORK_CONTEXT: { name: 'Work Context', questions: [2,3,4,5,6], scale: false },
  OVERALL_SAT: { name: 'Overall Satisfaction', questions: [7,8,9,10], scale: true },
  STEWARD_3A: { name: 'Steward Ratings', questions: [11-17], scale: true },
  STEWARD_3B: { name: 'Steward Access', questions: [19,20,21], scale: true },
  CHAPTER: { name: 'Chapter Effectiveness', questions: [22-26], scale: true },
  LEADERSHIP: { name: 'Local Leadership', questions: [27-32], scale: true },
  CONTRACT: { name: 'Contract Enforcement', questions: [33-36], scale: true },
  REPRESENTATION: { name: 'Representation Process', questions: [38-41], scale: true },
  COMMUNICATION: { name: 'Communication Quality', questions: [42-46], scale: true },
  MEMBER_VOICE: { name: 'Member Voice & Culture', questions: [47-51], scale: true },
  VALUE_ACTION: { name: 'Value & Collective Action', questions: [52-56], scale: true },
  SCHEDULING: { name: 'Scheduling/Office Days', questions: [57-63], scale: true },
  PRIORITIES: { name: 'Priorities & Close', questions: [65-68], scale: false }
};
```

### FEEDBACK_COLS (11 columns: A-K)

```javascript
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
```

---

## Sheet Structure (8 Visible + 6 Hidden)

### Core Data Sheets

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 1 | Config | Data | Master dropdown lists for validation (43 columns) |
| 2 | Member Directory | Data | All member data (31 columns) |
| 3 | Grievance Log | Data | All grievance cases (34 columns) |

### Dashboard Sheets

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 4 | 💼 Dashboard | View | Executive metrics dashboard with 12 analytics sections |
| 5 | 📊 Dashboard | View | Customizable metrics with dropdowns, My Cases tab for stewards |

### Tracking Sheets

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 6 | 📊 Member Satisfaction | Data | 68-question Google Form survey with dashboard, charts (82 cols + dashboard) |
| 7 | 💡 Feedback & Development | Data | Bug/feature tracking with priority (11 columns) |
| 8 | ✅ Function Checklist | Reference | Function reference guide organized by 14 phases |

### Help & Documentation Sheets

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 9 | 📚 Getting Started | Help | Step-by-step setup instructions, member/grievance guides, menu reference |
| 10 | ❓ FAQ | Help | Frequently asked questions organized by category (20+ Q&As) |
| 11 | 📖 Config Guide | Help | How to use Config tab: add/edit dropdowns, column reference, tips & warnings |

#### 💼 Dashboard - 12 Live Analytics Sections

| # | Section | Color | Metrics | Data Source |
|---|---------|-------|---------|-------------|
| 1 | QUICK STATS | 🟢 Green | Total Members, Active Stewards, Active Grievances, Win Rate, Overdue, Due This Week | `_Dashboard_Calc` |
| 2 | MEMBER METRICS | 🔵 Blue | Total Members, Active Stewards, Avg Open Rate, YTD Vol Hours | Member Directory |
| 3 | GRIEVANCE METRICS | 🟠 Orange | Open, Pending Info, Settled, Won, Denied, Withdrawn | Grievance Log |
| 4 | TIMELINE & PERFORMANCE | 🟣 Purple | Avg Days Open, Filed This Month, Closed This Month, Avg Resolution | Grievance Log |
| 5 | TYPE ANALYSIS | 🔷 Indigo | 5 issue categories × (Total, Open, Resolved, Win Rate, Avg Days) | Grievance Log |
| 6 | LOCATION BREAKDOWN | 🔵 Cyan | 5 locations × (Members, Grievances, Open Cases, Win Rate) | Config + Member Dir + Grievance Log |
| 7 | MONTH-OVER-MONTH TRENDS | 🔴 Red | Filed/Closed/Won × (This Month, Last Month, Change, % Change, Trend) | Grievance Log |
| 8 | STATUS LEGEND | ⬜ Gray | Color/icon reference guide | Static |
| 9 | STEWARD PERFORMANCE SUMMARY | 🟣 Purple | Total Stewards, Active w/Cases, Avg Cases, Vol Hours, Contacts | Member Dir + Grievance Log |
| 10 | TOP 30 BUSIEST STEWARDS | 🔴 Dark Red | Rank, Steward Name, Active Cases, Open, Pending Info, Total Ever | Grievance Log |
| 11 | TOP 10 PERFORMERS BY SCORE | 🟢 Green | Rank, Steward, Score, Win Rate %, Avg Days, Overdue | `_Steward_Performance_Calc` |
| 12 | STEWARDS NEEDING SUPPORT | 🔴 Red | Rank, Steward, Score, Win Rate %, Avg Days, Overdue | `_Steward_Performance_Calc` |

**All sections use live COUNTIF/COUNTIFS/AVERAGEIFS formulas that auto-update when source data changes.**

### Hidden Calculation Sheets

| # | Sheet Name | Purpose |
|---|------------|---------|
| 1 | `_Grievance_Calc` | Grievance → Member Directory sync (AB-AD) |
| 2 | `_Grievance_Formulas` | Member → Grievance Log sync (C-D, X-AA) |
| 3 | `_Member_Lookup` | Member data lookup formulas |
| 4 | `_Steward_Contact_Calc` | Steward contact tracking (Y-AA) |
| 5 | `_Dashboard_Calc` | Dashboard summary metrics (15 key KPIs) |
| 6 | `_Steward_Performance_Calc` | Per-steward performance scores with weighted formula |

---

## Config Sheet & Dropdown Validations

### Config Columns (15 columns + HOME_TOWNS at AF + Strategic Command at AS-AW)

| Column | Name | Used By |
|--------|------|---------|
| A | Job Titles | Member Directory (D) |
| B | Office Locations | Member Directory (E), Grievance Log (Z) |
| C | Units | Member Directory (F), Grievance Log (Y) |
| D | Office Days | Member Directory (G) |
| E | Yes/No | Member Directory (N, U, V, W) |
| F | Supervisors | Member Directory (L) |
| G | Managers | Member Directory (M) |
| H | Stewards | Member Directory (P, Z), Grievance Log (AA) |
| I | Grievance Status | Grievance Log (E) |
| J | Grievance Step | Grievance Log (F) |
| K | Issue Category | Grievance Log (W) |
| L | Articles Violated | Grievance Log (V) |
| M | Communication Methods | Member Directory (J) |
| N | (blank) | - |
| O | Grievance Coordinators | Admin use |
| AF | Home Towns | Member Directory (X) |
| **AS** | **Chief Steward Email** | **Strategic Command Center (escalation alerts)** |
| **AT** | **Unit Codes** | **Strategic Command Center (ID generation)** |
| **AU** | **Archive Folder ID** | **Strategic Command Center (PDF archive)** |
| **AV** | **Escalation Statuses** | **Strategic Command Center (alert triggers)** |
| **AW** | **Escalation Steps** | **Strategic Command Center (alert triggers)** |

### Member Directory Dropdowns (16 columns)

| Column | Field | Config Source | Multi-Select |
|--------|-------|---------------|--------------|
| D | Job Title | JOB_TITLES (A) | No |
| E | Work Location | OFFICE_LOCATIONS (B) | No |
| F | Unit | UNITS (C) | No |
| G | Office Days | OFFICE_DAYS (D) | **Yes** |
| J | Preferred Communication | COMM_METHODS (N) | **Yes** |
| K | Best Time to Contact | BEST_TIMES (AE) | **Yes** |
| L | Supervisor | SUPERVISORS (F) | No |
| M | Manager | MANAGERS (G) | No |
| N | Is Steward | YES_NO (E) | No |
| O | Committees | STEWARD_COMMITTEES (I) | **Yes** |
| P | Assigned Steward | STEWARDS (H) | **Yes** |
| U | Interest: Local | YES_NO (E) | No |
| V | Interest: Chapter | YES_NO (E) | No |
| W | Interest: Allied | YES_NO (E) | No |
| X | Home Town | HOME_TOWNS (AF) | No |
| Z | Contact Steward | STEWARDS (H) | No |

### Grievance Log Dropdowns (5 columns)

| Column | Field | Config Source |
|--------|-------|---------------|
| B | Member ID | Member Directory (A) - dynamic |
| E | Status | GRIEVANCE_STATUS (I) |
| F | Current Step | GRIEVANCE_STEP (J) |
| V | Articles Violated | ARTICLES (L) |
| W | Issue Category | ISSUE_CATEGORY (K) |

### Multi-Select Functionality

Columns marked as **Multi-Select** support comma-separated values for multiple selections.

**Auto-Open Mode (Recommended):**
1. Go to **🔧 Tools > ☑️ Multi-Select > ⚡ Enable Auto-Open**
2. Now clicking any multi-select cell automatically opens the dialog!
3. To disable: **🔧 Tools > ☑️ Multi-Select > 🚫 Disable Auto-Open**

**Manual Mode:**
1. Select a cell in a multi-select column (G, J, K, O, or P)
2. Go to **🔧 Tools > ☑️ Multi-Select > 📝 Open Editor**
3. Check multiple options in the dialog
4. Click **Save** to apply

**Storage format:** Values are stored as comma-separated text (e.g., "Monday, Wednesday, Friday")

**Validation:** Multi-select columns show a dropdown for convenience but accept any text value to allow multiple selections.

---

## Menu System (6 Menus)

The menu system has been reorganized from 9 menus to 6 logical groups (5 original + Strategic Command Center):

```
📊 509 Dashboard
├── 📊 Dashboard (with My Cases tab for stewards)
├── 📋 Dashboard Pend
├── 📊 Member Satisfaction (enhanced with trends & drill-down)
├── 📱 Mobile Dashboard
├── 🔍 Search Members
└── 📱 Get Mobile App URL

📋 Grievances
├── ➕ Start New Grievance
├── 📋 View Active Grievances
├── 📊 Sort by Status Priority
├── 🔄 Refresh Grievance Data
├── 🔄 Refresh Member Data
├── 📁 Drive Folders (submenu)
│   ├── 📁 Setup Folder for Grievance
│   ├── 📁 View Grievance Files
│   └── 📁 Batch Create All Folders
├── 📅 Calendar (submenu)
│   ├── 📅 Sync Deadlines to Calendar
│   ├── 📅 View Upcoming Deadlines
│   └── 🗑️ Clear Calendar Events
└── 📬 Notifications (submenu)
    ├── ⚙️ Notification Settings
    ├── ⚙️ Alert Settings
    ├── 📧 Send Steward Alerts Now
    └── 🧪 Test Notifications

👁️ View
├── 📅 Simplify Timeline (Hide Steps)
├── 📅 Show Full Timeline
├── ♿ Comfort View (submenu)
│   ├── ♿ Comfort View Panel
│   ├── 🎯 Focus Mode
│   ├── 🔲 Zebra Stripes
│   ├── 📝 Quick Capture
│   └── 🍅 Pomodoro
├── 🎨 Theming (submenu)
│   ├── 🎨 Theme Manager
│   ├── 🌙 Dark Mode
│   └── 🔄 Reset Theme
└── 🎨 Comfort View Setup (submenu)
    ├── 🎨 Setup Comfort View
    └── ↩️ Undo Comfort View

⚙️ Settings
├── 📊 Rebuild Dashboard
├── 🔄 Refresh All Formulas
├── ⚙️ Setup Data Validations
├── 🔧 REPAIR DASHBOARD
├── ☑️ Multi-Select (submenu)
│   ├── 📝 Open Editor
│   ├── ⚡ Enable Auto-Open
│   └── 🚫 Disable Auto-Open
├── 🔗 Live Formulas (submenu)
│   ├── 🔗 Setup Live Grievance Links
│   └── 👤 Clear Member ID Validation
├── ⚡ Triggers (submenu)
│   ├── ⚡ Install Auto-Sync Trigger
│   └── 🚫 Remove Auto-Sync Trigger
├── ✅ Validation (submenu)
│   ├── 🔍 Run Bulk Validation
│   ├── ⚙️ Validation Settings
│   ├── 🧹 Clear Validation Indicators
│   └── ⚡ Install Validation Trigger
└── 🎨 Comfort View Setup (submenu)

🔧 Admin
├── 🔍 DIAGNOSE SETUP
├── 🔍 Verify Hidden Sheets
├── 🔧 Hidden Sheets (submenu)
│   ├── 🔧 Setup All Hidden Sheets
│   └── 🔧 Repair All Hidden Sheets
├── 🔄 Data Sync (submenu)
│   ├── 🔄 Sync All Data Now
│   ├── 🔄 Sync Grievance → Members
│   └── 🔄 Sync Members → Grievances
├── 🧪 Testing (submenu)
│   ├── 🧪 Run All Tests
│   ├── ⚡ Quick Tests
│   └── 📊 View Test Results
├── 🗄️ Cache (submenu)
│   ├── 🗄️ Cache Status
│   ├── 🔥 Warm Caches
│   └── 🗑️ Clear Caches
└── 🎭 Demo Data (submenu) - Only visible if DEMO_MODE_DISABLED != 'true'
    ├── 🚀 Seed All Sample Data
    ├── ☢️ NUKE SEEDED DATA
    ├── 🧹 Clear Config Dropdowns Only
    └── 🔄 Restore Config & Dropdowns

📊 509 Command (Strategic Command Center v3.6.0)
├── 👁️ Executive Command (PII) - Internal dashboard with member names
├── 🫂 Member Analytics (No PII) - PII-safe dashboard
├── 📩 Send Member Dashboard Link
├── 🚀 Strategic Pro Moves (submenu)
│   ├── 🔥 Generate Unit Hot Zones
│   ├── 🌟 Identify Rising Stars
│   ├── 📉 Management Hostility Report
│   └── 📝 Bargaining Cheat Sheet
├── 🆔 ID & Data Engines (submenu)
│   ├── 🆔 Generate Missing Member IDs
│   ├── 🔍 Check Duplicate IDs
│   └── 📄 Create PDF for Selected Grievance
├── 👤 Steward Management (submenu)
│   ├── ⬆️ Promote to Steward
│   └── ⬇️ Demote Steward
├── 🎨 Styling & Theme (submenu)
│   ├── 🎨 Apply Global Styling
│   └── 🔄 Reset to Default Theme
└── ⚙️ Automation (submenu)
    ├── 🔄 Force Global Refresh
    ├── 🌙 Enable Midnight Auto-Refresh
    ├── ❌ Disable Midnight Auto-Refresh
    ├── 🔔 Enable 1AM Dashboard Refresh
    └── 📑 Email Weekly PDF Snapshot

    NOTE: Delete DeveloperTools.gs before production to remove all demo functions
```

---

## Config Sheet Columns

| Column | Name | Content |
|--------|------|---------|
| A | Job Titles | User populates |
| B | Office Locations | User populates |
| C | Units | User populates |
| D | Office Days | Monday-Sunday (preset) |
| E | Yes/No | Yes, No (preset) |
| F | Supervisors | User populates |
| G | Managers | User populates |
| H | Stewards | User populates |
| I | Steward Committees | User populates |
| J | Grievance Status | Open, Pending Info, Settled, etc. (preset) |
| K | Grievance Step | Informal, Step I, Step II, etc. (preset) |
| L | Issue Category | Discipline, Workload, etc. (preset) |
| M | Articles Violated | Art. 1 - Art. 26 (preset) |
| N | Communication Methods | Email, Phone, Text, In Person (preset) |
| O | Grievance Coordinators | User populates |
| **P** | **Grievance Form URL** | Auto-set via Save Form URLs to Config |
| **Q** | **Contact Form URL** | Auto-set via Save Form URLs to Config |
| AF | Home Towns | User populates |
| **AR** | **Satisfaction Survey URL** | Auto-set via Save Form URLs to Config |

---

## Color Scheme

```javascript
var COLORS = {
  PRIMARY_PURPLE: '#7C3AED',  // Main brand
  UNION_GREEN: '#059669',     // Success
  SOLIDARITY_RED: '#DC2626',  // Alert/urgent
  PRIMARY_BLUE: '#7EC8E3',    // Light blue
  ACCENT_ORANGE: '#F97316',   // Warnings
  LIGHT_GRAY: '#F3F4F6',      // Backgrounds
  TEXT_DARK: '#1F2937',       // Primary text
  WHITE: '#FFFFFF'
};
```

---

## Usage Patterns

### Dynamic Column References

```javascript
// CORRECT - Dynamic column reference
var statusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
var formula = '=COUNTIF(\'Grievance Log\'!' + statusCol + ':' + statusCol + ',"Open")';

// WRONG - Hardcoded column (NEVER DO THIS)
var formula = '=COUNTIF(\'Grievance Log\'!E:E,"Open")';
```

### Safe Row Writes

```javascript
// CORRECT - Protects header row
var startRow = Math.max(sheet.getLastRow() + 1, 2);

// WRONG - May overwrite headers
var startRow = sheet.getLastRow() + 1;
```

### Dynamic Sheet Names

```javascript
// CORRECT
var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

// WRONG
var sheet = ss.getSheetByName('Member Directory');
```

---

## Seed Data Limits

| Function | Max Count | Batch Size |
|----------|-----------|------------|
| SEED_MEMBERS() | 2,000 | 50 rows |
| SEED_GRIEVANCES() | 300 | 25 rows |

---

## Hidden Sheet Architecture (Self-Healing)

The system uses 6 hidden calculation sheets with auto-sync triggers for cross-sheet data population. Formulas are stored in hidden sheets and synced to visible sheets, making them **self-healing** - if formulas are accidentally deleted, running REPAIR_DASHBOARD() restores them.

### Hidden Sheets (6 total)

| Sheet | Source | Destination | Purpose |
|-------|--------|-------------|---------|
| `_Grievance_Calc` | Grievance Log | Member Directory | AB-AD (Has Open Grievance?, Status, Days to Deadline) |
| `_Grievance_Formulas` | Member Directory | Grievance Log | C-D (Name), H-P (Timeline), S-U (Days Open, Next Action, Days to Deadline), X-AA (Contact) |
| `_Member_Lookup` | Member Directory | Grievance Log | Member data lookup |
| `_Steward_Contact_Calc` | Member Directory | Contact Reports | Y-AA (Contact tracking) |
| `_Dashboard_Calc` | Both | 💼 Dashboard | 15 summary metrics (Win Rate, Overdue, Due This Week, etc.) |
| `_Steward_Performance_Calc` | Grievance Log | 💼 Dashboard | Per-steward performance scores with weighted formula |

### Auto-Sync Trigger

The `onEditAutoSync` trigger automatically syncs data when:
- Grievance Log is edited → Updates Member Directory columns AB-AD, then auto-sorts by status
- Member Directory is edited → Updates Grievance Log columns C-D, X-AA

### Key Functions (08_Code.gs - Hidden Sheet Management)

| Function | Purpose |
|----------|---------|
| `setupAllHiddenSheets()` | Create all 6 hidden sheets with formulas |
| `repairAllHiddenSheets()` | Recreate sheets, install trigger, sync data |
| `installAutoSyncTrigger()` | Install the onEdit auto-sync trigger |
| `verifyHiddenSheets()` | Verify all sheets and triggers are working |
| `syncAllData()` | Manual sync of all cross-sheet data |
| `syncGrievanceToMemberDirectory()` | Sync grievance data to members |
| `syncMemberToGrievanceLog()` | Sync member data to grievances |
| `sortGrievanceLogByStatus()` | Auto-sort by status priority and deadline |
| `setupDashboardCalcSheet()` | Create dashboard metrics calculation sheet |

### Self-Healing

If hidden sheets get corrupted or deleted:
1. Run `REPAIR_DASHBOARD()` from Setup menu
2. Or run `repairAllHiddenSheets()` from Administrator menu

This recreates all formulas and reinstalls the auto-sync trigger.

---

## Known Issues / Later TODO

### BUG: Days to Deadline Shows Duplicate Values

**Status:** FIXED (2025-12-16)
**Discovered:** 2025-12-16
**Severity:** Medium

**Issue:**
The "Days to Deadline" column (U) in the Grievance Log displayed identical values for multiple rows (e.g., `17.71857539` repeated for all grievances).

**Root Cause:**
The hidden sheet `_Grievance_Formulas` used ARRAYFORMULA with a FILTER-based row index. ARRAYFORMULA doesn't expand correctly when its source column is a FILTER result, causing all rows to receive the same calculated value.

**Fix Applied:**
Changed `syncGrievanceFormulasToLog()` in `HiddenSheets.gs` to calculate Days Open, Next Action Due, and Days to Deadline directly in JavaScript from the grievance row data, bypassing the problematic hidden sheet formulas.

**Calculations now performed directly:**
- **Days Open**: `(Date Closed or Today) - Date Filed` (in whole days)
- **Next Action Due**: Based on Current Step (Informal→Filing Deadline, Step I→Step I Due, etc.)
- **Days to Deadline**: `Next Action Due - Today` (in whole days)
- All deadline dates (Filing Deadline, Step I Due, etc.) also calculated directly

---

## Changelog

### Version 2.2.0 / v3.48 (2026-01-14) - Survey Verification, Data Integrity & Steward Workload

**New Survey Verification System:**
- Survey responses are now verified against the Member Directory
- Email matching: Survey respondent email is compared against member contact emails
- Verified responses are marked with `Verified: Yes`
- Unmatched emails are flagged as `Pending Review` for manual review
- Rejected submissions excluded from all statistics

**New Quarterly Tracking with History:**
- Each survey response is assigned a quarter (e.g., "2026-Q1")
- Members can submit multiple responses per quarter - only the latest counts in stats
- Historical responses preserved (not deleted) with `IS_LATEST: No` marking
- Previous responses marked with reference to newer submission (`SUPERSEDED_BY` column)
- Toggle in Member Dashboard to include/exclude historical responses from statistics

**New Flagged Submissions Review Interface:**
- Menu: 509 Dashboard > Survey Tools > 🔍 Review Flagged Submissions
- Shows count of pending review submissions
- Displays email addresses of unverified submissions (survey answers protected)
- One-click Approve or Reject actions for administrators
- Functions: `showFlaggedSubmissionsReview()`, `getFlaggedSubmissionsData()`, `approveFlaggedSubmission()`, `rejectFlaggedSubmission()`

**Public Member Dashboard Updates:**
- Survey statistics now filter to only show Verified responses
- Response rate calculated from unique verified member IDs
- New toggle: "Include historical responses" checkbox
- When enabled, shows all responses including superseded entries
- Visual warning when historical data is included

**New SATISFACTION_COLS Columns (Constants.gs):**
| Column | Name | Purpose |
|--------|------|---------|
| CE (83) | EMAIL | Email address from form submission |
| CF (84) | VERIFIED | Yes / Pending Review / Rejected |
| CG (85) | MATCHED_MEMBER_ID | Member ID if email matched |
| CH (86) | QUARTER | Quarter string (e.g., "2026-Q1") |
| CI (87) | IS_LATEST | Yes/No - Is this the latest for this member this quarter? |
| CJ (88) | SUPERSEDED_BY | Row number of newer response (if superseded) |
| CK (89) | REVIEWER_NOTES | Notes from reviewer |

**Email Collection Configuration:**
For verification to work, the Google Form must collect respondent emails. Two options:
1. **Enable "Collect email addresses"** in Google Form settings (recommended)
2. **Add an "Email Address" question** to the form

If neither is configured, all submissions will be marked as "Pending Review" and require manual approval.

**Updated Functions:**
- `onSatisfactionFormSubmit(e)` - Now includes email verification and quarterly tracking
- `validateMemberEmail(email)` - Validates email against Member Directory
- `getCurrentQuarter()` - Returns current quarter string
- `getPublicSurveyData(includeHistory)` - Now accepts toggle for historical data

**Files Modified (Survey Verification):**
- `Constants.gs` - Added SATISFACTION_COLS verification columns
- `Code.gs` - Updated form handler, added review interface, updated public dashboard

**New Data Integrity Module (DataIntegrity.gs):**
- New file: `DataIntegrity.gs` (~1000 lines) providing comprehensive data management
- Menu: 🛡️ Data Integrity - New top-level menu with integrity tools

**Batch Operations Utilities:**
- `batchSetValues()` - Write multiple values in single operation
- `batchSetRowValues()` - Update multiple columns in one row efficiently
- `batchAppendRows()` - Add multiple rows at once (replaces slow appendRow loops)

**Error Handling:**
- `executeWithRetry()` - Retry with exponential backoff for transient failures
- `safeSheetOperation()` - Wrapped operations with error capture

**Confirmation Dialogs:**
- `getOrCreateSheetSafe()` - Confirms before deleting sheets with data
- `confirmDestructiveAction()` - Reusable confirmation helper

**Dynamic Validation:**
- `setDropdownValidationDynamic()` - Uses getLastRow() instead of fixed 100
- `setMultiSelectValidationDynamic()` - Dynamic multi-select validation

**Duplicate ID Validation:**
- `checkDuplicateMemberId()` - Check for duplicate Member IDs
- `validateMemberIdOnEdit()` - Real-time duplicate detection on edit

**Ghost Validation (Orphaned Grievances):**
- `findOrphanedGrievances()` - Find grievances with invalid Member IDs
- `highlightOrphanedGrievances()` - Visual highlighting of issues
- `runScheduledGhostValidation()` - Email alerts for data issues

**Steward Load Balancing:**
- `calculateStewardWorkload()` - Calculate load scores per steward
- `showStewardWorkloadDashboard()` - Visual dashboard with metrics
- `getStewardWithLowestWorkload()` - For auto-assignment

**Self-Healing Config Tool:**
- `findMissingConfigValues()` - Scan for values not in Config dropdowns
- `showConfigHealthCheck()` - UI for reviewing issues
- `autoFixMissingConfigValues()` - Auto-add missing values

**Enhanced Audit Logging:**
- `logIntegrityEvent()` - Comprehensive logging with timestamps
- `logGrievanceStatusChange()` - Status change tracking
- `logStewardAssignmentChange()` - Assignment change tracking
- `showAuditLogViewer()` - UI to view recent entries

**Auto-Archive:**
- `archiveClosedGrievances()` - Move old closed cases to archive
- `showArchiveDialog()` - UI with configurable age threshold
- `restoreFromArchive()` - Restore archived grievances

**Visual Enhancements:**
- `applyDeadlineHeatmap()` - Conditional formatting for deadline urgency
- `createMobileStewardPortal()` - Mobile-optimized view for field stewards

**Enhanced Quick Actions (MobileQuickActions.gs):**
- 📧 Email Survey to Member - Send satisfaction survey link via email
- 📧 Email Contact Form - Send member info update form link
- 📧 Email Dashboard Link - Share spreadsheet access with members
- 📧 Email Grievance Status - Send status update for active grievances

**Files Modified (Data Integrity):**
- `DataIntegrity.gs` - New file with all data integrity functions
- `MobileQuickActions.gs` - Added email quick actions
- `build.js` - Added DataIntegrity.gs to module list and dependencies
- `ConsolidatedDashboard.gs` - Rebuilt with DataIntegrity module

---

### Version 2.1.0 (2026-01-14) - Dashboard Renaming, My Cases Tab & Member Satisfaction Enhancements

**Dashboard Naming Changes:**
- Renamed "Custom View Dashboard" to **"Dashboard"**
- Renamed "Smart Dashboard" to **"Dashboard Pend"**
- Updated menu items, dialog titles, and HTML headers across ConsolidatedDashboard.gs

**Removed Quick Actions Menu:**
- Removed "Quick Actions" menu item from 509 Dashboard menu (Quick Actions checkboxes in sheets still work)

**New My Cases Tab (for Stewards):**
- Added "My Cases" tab to the Dashboard modal popup
- Shows grievances assigned to the current steward
- Displays stats: Total Cases, Pending, In Progress
- Includes status filtering and expandable case details
- New backend function: `getMyStewardCases()`

**Enhanced Member Satisfaction Dashboard:**
- **New Trends Tab** with time period filtering (All Time, Year, 90 Days, 30 Days)
- **Line Chart** showing satisfaction score trends over time
- **Donut Chart** for visualizing top issues distribution
- **Renamed "Loyalty Score" to "Member Advocacy Index (MAI)"** with clearer explanation
- **Clickable Worksite Drill-down** - click any worksite bar to see detailed responses in a modal
- **Diverse Color Palette** - added teal, indigo, pink, amber accent colors
- **Clickable Response Count** - navigates to Responses tab
- **Response Count by Month** with rainbow-colored bars
- New backend functions: `getSatisfactionTrendData()`, `getSatisfactionLocationDrill()`, `getSheetLastRow()`

**Files Modified:**
- `ConsolidatedDashboard.gs` - All dashboard changes applied to deployment file
- `Code.gs`, `MobileQuickActions.gs` - Source file updates

---

### Version 2.0.6 (2026-01-13) - Fix Dashboard Bar Charts and Mobile WebApp

**Bug Fix: Bar charts showing 100% full in all dashboards**

All bar charts in Member Satisfaction Dashboard, Smart Dashboard, and Mobile Quick Actions were displaying at full width regardless of actual data values.

**Root Cause:**
CSS `width` property was using `%25` (URL-encoded percent sign) instead of `%`. CSS does not URL-decode inline style values, so `width:70%25` was invalid and browsers defaulted to full width or ignored the value entirely.

**Example of bug:**
```javascript
// BEFORE (broken) - %25 is not valid CSS
style="width:"+pct+"%25;background:#059669"

// AFTER (fixed) - uses proper % symbol
style="width:"+pct+"%;background:#059669"
```

**Files Fixed:**
- `Code.gs` - 4 bar chart instances in Member Satisfaction Dashboard (lines 6915, 6963, 6973, 6996)
- `MobileQuickActions.gs` - 6 bar chart instances in Smart Dashboard (lines 1166, 1174, 1182, 1183, 1184, 1191)
- `ConsolidatedDashboard.gs` - All instances updated via rebuild

**Affected Dashboards:**
- Member Satisfaction Dashboard (📊 Member Satisfaction modal)
  - By Section satisfaction scores
  - Satisfaction by Worksite
  - Satisfaction by Role
  - Top Member Priorities
- Smart Dashboard / Mobile Dashboard (📊 509 Dashboard modal)
  - Members by Location
  - Members by Unit
  - Grievance Status Distribution
  - Top Issue Categories

**Enhancement: Mobile WebApp Error Handling**

Improved error handling and debugging for the standalone Mobile Web App (`WebApp.gs`):

- Added `try-catch` error handling to `getWebAppGrievanceList()` and `getWebAppMemberList()`
- Added `Logger.log()` statements for server-side debugging
- Added `console.log()` statements for client-side debugging in browser dev tools
- Error messages now display the actual error instead of generic "Error loading data"

**Important Deployment Note:**
After updating the code in Google Apps Script, you must **redeploy the web app** for changes to take effect:
1. Go to Extensions → Apps Script
2. Click "Deploy" → "Manage deployments"
3. Edit existing deployment or create new one
4. Click "Deploy"

---

### Version 2.0.5 (2026-01-13) - Fix Grievance Seeding Range Error

**Bug Fix: "Coordinates outside dimensions" error when seeding grievances**

The error occurred because `SEED_GRIEVANCES` and `SEED_MEMBERS` functions tried to write data to columns beyond the sheet's default dimensions. When a Google Sheet is created via `insertSheet()`, it only has 26 columns (A-Z), but the Grievance Log requires 35 columns (A-AI) and Member Directory requires 32 columns (A-AF).

**Root Cause:**
- `SEED_GRIEVANCES` at line 1111 tried to write 34 columns of data
- If Grievance Log sheet only had 26 columns, `getRange()` would fail

**Fix Applied:**
- Added `ensureMinimumColumns(sheet, requiredColumns)` helper function
- Called at start of `SEED_GRIEVANCES`, `SEED_MEMBERS`, and `SEED_MEMBERS_ONLY`
- Automatically expands sheet columns if needed before writing data

**Files Updated:**
- `DeveloperTools.gs` - Added helper function and column checks
- `ConsolidatedDashboard.gs` - Same updates for deployed version

---

### Version 2.0.4 (2026-01-13) - Mobile Web App v2.0 Enhancements

**Major upgrade to the standalone Mobile Web App (WebApp.gs) with full feature parity**

---

#### 1. Dashboard (Home) Page Enhanced

**6 Clickable Stats in 3x2 Grid:**
- Members - links to Members page
- Grievances - links to Grievances page
- Active - links to Grievances filtered to Open
- Pending - links to Grievances filtered to Pending
- Overdue - links to Grievances filtered to Overdue
- Win Rate - displays grievance success percentage

**Overdue Preview Section:**
- Shows top 3 overdue cases dynamically
- Red danger styling with "View All X Overdue Cases" button
- Links directly to filtered grievance view

**New Functions:**
- `getWebAppDashboardStats()` - Returns stats with winRate and totalMembers

---

#### 2. Grievances (Cases) Page Enhanced

**New Overdue Filter Button:**
- Added "⚠️ Overdue" filter pill with danger (red) styling
- Pulsing animation on overdue badges

**Expandable Card Details:**
- Tap any grievance card to expand/collapse details
- Shows: Filed date, Incident date, Next due, Days open, Location, Articles, Steward, Resolution
- Overdue cards have red left border

**URL Filter Parameter:**
- Supports `?page=grievances&filter=overdue` for direct linking
- Auto-selects filter pill based on URL parameter

**Updated Data Function:**
- `getWebAppGrievanceList()` now returns 16 fields (same as Interactive Dashboard)
- Includes: isOverdue, daysToDeadline, daysOpen, incidentDate, nextActionDue, location, articles, steward, resolution

---

#### 3. New Members Page

**New page accessible via `?page=members`**

**Features:**
- Search by name, ID, title, location
- Filter pills: All, Stewards, With Grievance
- Expandable cards with email, phone, unit, supervisor
- Steward badge and Open Grievance badge
- Orange left border for members with open grievances

**New Functions:**
- `getWebAppMemberListHtml()` - Returns Members page HTML
- `getWebAppMemberList()` - Returns member data (12 fields)

---

#### 4. New Links Page

**New page accessible via `?page=links`**

**Features:**
- Forms section: Grievance Form, Contact Form, Satisfaction Survey (if configured)
- Resources section: Spreadsheet link, GitHub Repository link
- GitHub card with dark styling

**New Functions:**
- `getWebAppLinksHtml()` - Returns Links page HTML
- `getWebAppResourceLinks()` - Returns configured form URLs and GitHub link

---

#### 5. Navigation Enhanced

**5-Item Bottom Navigation (all pages):**
- Home (Dashboard)
- Search
- Cases (Grievances)
- Members
- Links

**Updated styling for 5-item layout:**
- Smaller icons (22px)
- Compact padding
- Shorter labels

---

**Files Changed:**
- `WebApp.gs`: Complete rewrite with 5 pages and enhanced features
- `README.md`: Updated Mobile Web App section to v2.0

---

### Version 2.0.3 (2026-01-13) - Member Add/Modify Forms & Chart Fixes

**Added member ADD/MODIFY functionality to Interactive Dashboard and fixed By Section charts**

**CRITICAL FIX: ConsolidatedDashboard.gs fully synced with MobileQuickActions.gs**

> **Note:** Per the deployment model, only `ConsolidatedDashboard.gs` is deployed.
> All feature development occurs in `MobileQuickActions.gs` and must be synced to `ConsolidatedDashboard.gs`.
> This version includes a complete sync of all Interactive Dashboard features.

---

#### 1. Member ADD/MODIFY Forms in Interactive Dashboard

**New Member Add Form:**
- Added "➕ Add New Member" button to Members tab
- Modal form with fields: First Name, Last Name, Job Title, Email, Phone, Work Location, Unit, Office Days (multi-select), Supervisor, Is Steward
- Auto-generates Member ID using `generateNameBasedId()` pattern (M + first 2 chars + last 2 chars + 3 digits)
- Saves new member to Member Directory sheet

**Member Edit Form:**
- Added "✏️ Edit Member" button on each expanded member item
- Pre-populates form with existing member data
- Updates member data in place without changing Member ID

**Server-side Function:**
- `saveInteractiveMember(memberData, mode)` - Handles both add and edit operations
- Validates required fields (first name, last name)
- Returns success status with member ID

**Files Changed:**
- `MobileQuickActions.gs`: Lines 765-792 (Add Member form modal HTML)
- `MobileQuickActions.gs`: Lines 923 (Edit button in member list)
- `MobileQuickActions.gs`: Lines 963-1055 (JavaScript form functions)
- `MobileQuickActions.gs`: Lines 1668-1745 (saveInteractiveMember server function)

---

#### 2. Office Days Filter Added

**New Filter Dropdown:**
- Added "All Office Days" filter dropdown to Members tab
- Filters members by their office days (Monday, Tuesday, etc.)
- Days sorted in weekday order
- Works in combination with Location/Unit filters and search

**Files Changed:**
- `MobileQuickActions.gs`: Lines 891-908 (loadMemberFilters updated)
- `MobileQuickActions.gs`: Lines 911-912 (resetMemberFilters updated)
- `MobileQuickActions.gs`: Lines 950 (filterMembers updated)

---

#### 3. GitHub Repository Link Added

**Links Tab Enhancement:**
- Added "📦 GitHub Repository" link to External Links section
- Links to https://github.com/Woop91/MULTIPLE-SCRIPS-REPO

**Files Changed:**
- `MobileQuickActions.gs`: Lines 1260-1262 (renderResources updated)

---

#### 4. By Section Chart Fixes (Member Satisfaction Dashboard)

**Fixed 100% Bar Issue:**
- Charts now skip sections with zero responses instead of showing full bars
- Added `hasValidData` check to show "No survey responses yet" message when appropriate
- Bars properly clamp to 0-100% range

**Removed Redundant Detail Cards:**
- Replaced repetitive "Section Details" cards with actionable insights
- Now shows "Areas Needing Attention" (scores < 6) and "Strong Performance" (scores >= 8)
- More useful summary instead of repeating the same information from the bar chart

**Added Clarifying Labels:**
- Chart title now shows "(1-10 Scale)" for clarity
- Added subtitle "Sorted by score - areas needing attention shown first"
- Bar values now show "X responses" instead of just a number

**Files Changed:**
- `Code.gs`: Lines 6899-6937 (renderSections function completely rewritten)
- `ConsolidatedDashboard.gs`: Lines 7713-7751 (same changes mirrored)
- `ConsolidatedDashboard.gs`: Lines 7777, 7787, 7795-7796 (n= labels changed to responses/members)

---

#### 5. ConsolidatedDashboard.gs Full Sync

**Critical sync to ensure deployed file has all features from source files.**

**Data Functions Updated:**
- `getInteractiveMemberData()` - Now returns all 16 member fields (firstName, lastName, email, phone, officeDays, unit, supervisor, hasOpenGrievance, assignedSteward, etc.)
- `getInteractiveGrievanceData()` - Now returns all 16 grievance fields (currentStep, isOverdue, daysToDeadline, daysOpen, incidentDate, nextActionDue, location, articles, steward, resolution, etc.)
- Both functions now properly skip blank rows (validates ID starts with M/G)

**Grievances Tab Enhanced:**
- Added "⚠️ Overdue" filter button with danger styling
- Added expandable grievance details with all fields on click
- Added `toggleGrievanceDetail()`, `showGrievanceDetail()` functions
- Filter buttons now highlight when active
- Proper Overdue filter that checks `isOverdue` property

**Overview Tab Enhanced:**
- Clickable stat cards (Total Members, Total Grievances, Open Cases)
- Added `showOpenCases()` function to jump to filtered grievances
- Added `loadOverduePreview()` to show overdue cases on overview

**CSS Styles Added:**
- `.badge-overdue` - Pulsing red badge for overdue items
- `.action-btn-danger` - Red button style for danger actions

**Files Changed:**
- `ConsolidatedDashboard.gs`: Lines 12182-12183 (badge-overdue CSS)
- `ConsolidatedDashboard.gs`: Lines 12192-12193 (action-btn-danger CSS)
- `ConsolidatedDashboard.gs`: Lines 12332-12338 (Grievances tab with Overdue filter)
- `ConsolidatedDashboard.gs`: Lines 12373-12415 (loadOverview, renderOverview, showOpenCases, loadOverduePreview)
- `ConsolidatedDashboard.gs`: Lines 12525-12624 (grievance functions with expandable details)
- `ConsolidatedDashboard.gs`: Lines 12750-12824 (getInteractiveMemberData, getInteractiveGrievanceData)

---

### Version 2.0.2 (2026-01-13) - Interactive Dashboard & Satisfaction Survey Improvements

**Comprehensive updates to the Custom View popup modal and Member Satisfaction Dashboard**

---

#### 1. Interactive Dashboard (Custom View) Enhancements

**Stability Improvements:**
- Added error handling with `safeRun()` wrapper for all JavaScript functions
- Added `.withFailureHandler()` to all `google.script.run` calls
- Error states now display friendly messages instead of blank pages

**Overdue Cases Filter:**
- Added **"⚠️ Overdue"** filter button to Grievances tab
- Overdue cases now shown on Overview tab with preview section
- Clicking "View All Overdue Cases" navigates to filtered view

**Blank Row Filtering:**
- All data functions now validate IDs start with "M" (members) or "G" (grievances)
- `getInteractiveOverviewData()`, `getInteractiveMemberData()`, `getInteractiveGrievanceData()`, `getInteractiveAnalyticsData()` all updated
- Prevents blank spreadsheet rows from appearing in modal

**Clickable List Items with Details:**
- Members: Click to expand showing email, phone, office days, supervisor, assigned steward
- Grievances: Click to expand showing incident date, next due date, days open, articles, resolution
- "Quick Actions" and "View in Sheet" buttons on expanded items

**Member Filters:**
- Location dropdown filter
- Unit dropdown filter
- Reset button to clear all filters
- Filters work in combination with search

**Resource Links Tab:**
- New "🔗 Links" tab added
- Shows Grievance Form, Contact Form, Satisfaction Survey links from Config sheet
- Quick access to open full spreadsheet
- Quick action buttons for common operations

**Navigation Functions:**
- `navigateToMemberInSheet(memberId)` - Jump to member in sheet
- `navigateToGrievanceInSheet(grievanceId)` - Jump to grievance in sheet
- `showMemberDirectory()`, `showGrievanceLog()`, `showConfigSheet()` - Tab navigation

**Files Changed:**
- `MobileQuickActions.gs`: Lines 594-1533 (complete overhaul)

---

#### 2. Member Satisfaction Dashboard Improvements

**NPS Terminology Changed to Intuitive Language:**
- "NPS Score" → "Member Advocacy Index" (formerly "Loyalty Score")
- "Strong NPS Score" → "Members Highly Recommend"
- "NPS Needs Improvement" → "Member Loyalty Needs Attention"
- Added "Moderate Member Loyalty" insight for scores 0-49

**Member Advocacy Index Explanation:**
- Added info card explaining Member Advocacy Index (MAI) meaning and score ranges
- Shows score ranges: 50+ = Excellent, 0-49 = Good, Below 0 = Needs work
- Explains it's based on "Would Recommend" question

**Clickable Response Details:**
- Responses tab items now expandable on click
- Shows individual scores: Satisfaction, Trust, Feel Protected, Would Recommend
- Shows Steward Contact status and Steward Rating if applicable

**Sample Size Labels Improved:**
- Changed "n=X" to "X responses" or "X members" throughout
- More intuitive for non-technical users

**Files Changed:**
- `Code.gs`: Lines 6543-7260 (Satisfaction Dashboard HTML/JS)
- `Code.gs`: Lines 7112-7133 (NPS insights)
- `Code.gs`: Lines 7179-7258 (Response data function expanded)
- `ConsolidatedDashboard.gs`: Same changes mirrored

---

### Version 1.9.1 (2026-01-13) - Grievance Log Bug Fixes & Quick Actions Checkbox

**Bug Fixes: Member ID Dropdown, Overdue Cases, Blank Row Counting + New Quick Actions Checkbox Feature**

Fixed three issues in the Grievance Log and Dashboard:

---

#### 1. Member ID Dropdown Removed

**Issue:** Member ID column (B) in Grievance Log had a dropdown validation that restricted entries to existing Member IDs from Member Directory.

**Fix:**
- Removed dropdown validation from Member ID column
- `setupGrievanceMemberDropdown()` now CLEARS validation instead of adding it
- Member ID now allows free text entry for flexibility

**Files Changed:**
- `Code.gs`: Commented out `setMemberIdValidation()` call in `setupDataValidations()`
- `Code.gs`: Updated `setupGrievanceMemberDropdown()` to clear validations
- `ConsolidatedDashboard.gs`: Same changes

---

#### 2. Overdue Cases Not Populating in Dashboard

**Issue:** The Dashboard "Overdue Cases" metric was always showing 0, even when overdue grievances existed.

**Root Cause:**
- Days to Deadline column stores the text `"Overdue"` for past-due cases (not negative numbers)
- `computeDashboardMetrics_()` only checked `typeof daysToDeadline === 'number'`
- String "Overdue" failed the number check, so overdue cases were never counted

**Fix:** Added explicit check for the string "Overdue" before the number check:
```javascript
// Before (broken):
if (typeof daysToDeadline === 'number') {
  if (daysToDeadline < 0) metrics.overdueCases++;
}

// After (fixed):
if (daysToDeadline === 'Overdue') {
  metrics.overdueCases++;
} else if (typeof daysToDeadline === 'number') {
  if (daysToDeadline < 0) metrics.overdueCases++;
}
```

**Files Changed:**
- `HiddenSheets.gs`: Line ~1744 in `computeDashboardMetrics_()`
- `ConsolidatedDashboard.gs`: Line ~10049 in `computeDashboardMetrics_()`

---

#### 3. Blank Rows Being Counted as Grievances

**Issue:** Adding blank rows to the Grievance Log caused them to be counted in the "Total Grievances" metric.

**Root Cause:** The formula used `COUNTA(...)-1` which counts any non-empty cell (including spaces or formatting artifacts).

**Fix:** Changed to `COUNTIF(...,"G*")` for grievance IDs and `COUNTIF(...,"M*")` for member IDs, which only counts cells starting with the valid ID prefix.

```javascript
// Before (broken):
['Total Grievances', '=COUNTA(...)-1']
['Total Members', '=COUNTA(...)-1']

// After (fixed):
['Total Grievances', '=COUNTIF(...,"G*")']
['Total Members', '=COUNTIF(...,"M*")']
```

**Files Changed:**
- `HiddenSheets.gs`: Lines ~1373-1375 in `setupDashboardCalcSheet()`
- `ConsolidatedDashboard.gs`: Lines ~2169-2171 and ~9686-9688

---

#### Additional Change: Days to Deadline Number Format

Changed Days to Deadline column format from `'0'` to `'General'` to better preserve the "Overdue" text display.

**Files Changed:**
- `HiddenSheets.gs`: Line ~610 in `syncGrievanceFormulasToLog()`
- `ConsolidatedDashboard.gs`: Line ~8924

---

#### 4. Quick Actions Checkbox (New Feature)

**Feature:** Added "⚡ Actions" checkbox column to both Grievance Log and Member Directory that opens the Quick Actions dialog when checked.

**Benefits:**
- No need to navigate to menu items - just check the checkbox in the row
- Checkbox auto-unchecks after opening the dialog so it can be reused
- Works with the existing Quick Actions dialogs (calendar sync, drive folder, email, etc.)

**Implementation:**
- Member Directory: Column AF (QUICK_ACTIONS = 32)
- Grievance Log: Column AI (QUICK_ACTIONS = 35)

**How It Works:**
1. User checks the ⚡ Actions checkbox in any data row
2. `onEditAutoSync()` trigger detects the checkbox change
3. Checkbox is immediately unchecked (reset for reuse)
4. Quick Actions dialog opens showing available actions for that row:
   - **Member Directory**: Start Grievance, Send Email, View Grievance History, Copy ID
   - **Grievance Log**: Sync to Calendar, Setup Drive Folder, Quick Status Update, Copy ID

**Files Changed:**
- `Constants.gs`: Added `QUICK_ACTIONS` to MEMBER_COLS and GRIEVANCE_COLS
- `Code.gs`: Updated `createMemberDirectory()` and `createGrievanceLog()` to add checkboxes
- `HiddenSheets.gs`: Updated `onEditAutoSync()` to handle Quick Actions checkbox clicks
- `ConsolidatedDashboard.gs`: Same updates

---

### Version 1.9.0 (2026-01-11) - Visual Enhancements, Progress Tracking & Grievance Form Workflow

**New Features: Data Validation, Heatmaps, Progress Bar & Automated Grievance Workflow**

Added visual data quality indicators, deadline heatmaps, grievance progress tracking, and automated grievance form workflow with Drive folder creation.

---

#### Member Directory Enhancements

**1. Empty Field Validation (Red Background)**
- Email field (Column H): Red background (`#ffcdd2`) when empty but Member ID exists
- Phone field (Column I): Red background (`#ffcdd2`) when empty but Member ID exists
- Formula: `=AND($A2<>"",ISBLANK($H2))` ensures only rows with members are highlighted

**2. Days to Deadline Heatmap (Column AD)**
| Days Remaining | Background | Text Color | Style |
|----------------|------------|------------|-------|
| Overdue or ≤0  | `#ffebee` (red) | `#c62828` | Bold |
| 1-3 days       | `#fff3e0` (orange) | `#e65100` | Bold |
| 4-7 days       | `#fffde7` (yellow) | `#f57f17` | Normal |
| 8+ days        | `#e8f5e9` (green) | `#2e7d32` | Normal |

**3. Column Sorting via Filter**
- Added filter row to Member Directory header
- All columns now sortable via dropdown (A-Z, Z-A)

---

#### Grievance Log Enhancements

**4. Days to Deadline Heatmap (Column U)**
Same color scheme as Member Directory - automatically applied when sheet is created.

**5. Grievance Progress Bar (Columns J-R)**
Visual progress indicator showing grievance stage via colored backgrounds:

| Current Step | Columns Highlighted | Color |
|--------------|---------------------|-------|
| Informal | None | Gray `#fafafa` |
| Step I | J-K | Soft blue `#e3f2fd` |
| Step II | J-O | Soft blue `#e3f2fd` |
| Step III | J-Q | Soft blue `#e3f2fd` |
| Closed/Won/Denied/Settled/Withdrawn | J-R | Soft green `#e8f5e9` |

- Progress bar spans: Step I Due → Step I Rcvd → Step II Appeal Due → Step II Appeal Filed → Step II Due → Step II Rcvd → Step III Appeal Due → Step III Appeal Filed → Date Closed
- Columns not yet reached remain light gray
- Completed grievances show all columns in green

**6. Auto-Sort Confirmation**
Grievance Log entries automatically sort by status priority (active cases first) and deadline urgency when edited.

---

#### Grievance Form Workflow (New)

**7. Pre-filled Google Form Integration**
- `startNewGrievance()`: Opens pre-filled Google Form with member data
- Form fields auto-populated from Member Directory (Member ID, name, job title, location, email, etc.)
- Steward info auto-populated from current user's session (if they're a steward in Member Directory)
- Default values: Date Filed = today, Step = I

**8. Automatic Form Submission Processing**
- `onGrievanceFormSubmit(e)`: Trigger handler for form submissions
- Generates unique Grievance ID (format: GXXXX123 based on member name)
- Adds grievance to Grievance Log with all form data
- Calculates deadlines via hidden sheet formulas
- Updates Member Directory grievance status

**9. Automatic Drive Folder Creation**
- Creates folder in "509 Dashboard - Grievance Files" root folder
- Folder name format: `GXXXX123 - FirstName LastName (MemberID)`
- Creates subfolders: 📄 Documents, 📧 Correspondence, 📝 Notes
- Automatically shares with Grievance Coordinators from Config (column O)
- Stores folder ID and URL in Grievance Log (columns AG, AH)

**10. Easy Trigger Setup**
- New menu: 👤 Dashboard > 📋 Grievance Tools > 📋 Setup Grievance Form Trigger
- Prompts for Google Form edit URL
- Creates installable trigger for form submissions

---

#### Personal Contact Info Form (New)

**11. Blank Form for Member Self-Registration**
- `sendContactInfoForm()`: Shows form link to share with members (open or copy)
- Form is blank - members fill out their own contact information
- Form fields: First/Last Name, Job Title, Unit, Work Location, Office Days, Communication Preferences, Best Time, Supervisor, Manager, Email, Phone, Interest levels

**12. Automatic Member Directory Updates**
- `onContactFormSubmit(e)`: Trigger handler for contact form submissions
- Matches member by First Name + Last Name
- **Existing members**: Updates all submitted fields in Member Directory
- **New members**: Creates new row with auto-generated Member ID (MXXXX123 format)
- Handles multi-select fields (Office Days, Preferred Communication, Best Time)

**13. Easy Trigger Setup**
- New menu: 👤 Dashboard > 👤 Member Tools > 📋 Get Contact Info Form Link
- New menu: 👤 Dashboard > 👤 Member Tools > ⚙️ Setup Contact Form Trigger
- Prompts for Google Form edit URL
- Creates installable trigger for contact form submissions

---

#### Member Satisfaction Survey Form (New)

**14. Survey Link Distribution**
- `getSatisfactionSurveyLink()`: Shows survey link to share with members (open or copy)
- Survey is blank - members fill out 68 questions about union satisfaction

**15. Automatic Response Recording**
- `onSatisfactionFormSubmit(e)`: Trigger handler for survey submissions
- Writes all 68 question responses to 📊 Member Satisfaction sheet
- Maps questions to SATISFACTION_COLS (A-BQ columns)
- Sections: Work Context, Overall Satisfaction, Steward Ratings, Steward Access, Chapter Effectiveness, Local Leadership, Contract Enforcement, Representation Process, Communication Quality, Member Voice & Culture, Value & Collective Action, Scheduling, Priorities

**16. Easy Trigger Setup**
- New menu: 👤 Dashboard > 📊 Survey Tools > 📊 Get Satisfaction Survey Link
- New menu: 👤 Dashboard > 📊 Survey Tools > ⚙️ Setup Survey Form Trigger
- Prompts for Google Form edit URL
- Creates installable trigger for survey submissions

---

**Code Changes:**

*Member Directory (`createMemberDirectory()` lines 497-576):*
- Lines 497-515: Empty Email/Phone validation rules
- Lines 517-554: Days to Deadline heatmap rules
- Lines 560-576: Filter for column sorting

*Grievance Log (`createGrievanceLog()` lines 645-739):*
- Lines 645-682: Days to Deadline heatmap rules
- Lines 684-731: Progress bar conditional formatting rules
- Lines 733-739: Apply all rules

*Grievance Form Workflow (Code.gs lines 3254-3813):*
- Lines 3258-3283: `GRIEVANCE_FORM_CONFIG` with form URL and 18 field entry IDs
- Lines 3290-3383: `startNewGrievance()` - opens pre-filled form
- Lines 3389-3412: `getCurrentStewardInfo_()` - get steward from session
- Lines 3419-3449: `buildGrievanceFormUrl_()` - build pre-filled URL
- Lines 3465-3561: `onGrievanceFormSubmit(e)` - form submission handler
- Lines 3568-3609: Helper functions (getFormValue_, parseFormDate_, getExistingGrievanceIds_)
- Lines 3615-3653: `createGrievanceFolderFromData_()` - create folder with subfolders
- Lines 3660-3683: `shareWithCoordinators_()` - share folder with coordinators
- Lines 3690-3780: `setupGrievanceFormTrigger()` - menu-driven trigger setup
- Lines 3787-3813: `testGrievanceFormSubmission()` - test function

*Contact Info Form Workflow (Code.gs):*
- `CONTACT_FORM_CONFIG` with form URL and 15 field entry IDs
- `sendContactInfoForm()` - shows blank form link (open or copy)
- `onContactFormSubmit(e)` - form submission handler (creates new or updates existing member)
- `getFormMultiValue_()` - helper for checkbox responses
- `setupContactFormTrigger()` - menu-driven trigger setup

*Satisfaction Survey Form Workflow (Code.gs):*
- `SATISFACTION_FORM_CONFIG` with survey form URL
- `getSatisfactionSurveyLink()` - shows survey link (open or copy)
- `onSatisfactionFormSubmit(e)` - form submission handler (writes to Member Satisfaction sheet)
- `setupSatisfactionFormTrigger()` - menu-driven trigger setup

*Menu Update (Code.gs):*
- Added "📋 Setup Grievance Form Trigger" to Grievance Tools submenu
- Added new "👤 Member Tools" submenu with:
  - "📋 Get Contact Info Form Link"
  - "⚙️ Setup Contact Form Trigger"
- Added new "📊 Survey Tools" submenu with:
  - "📊 Get Satisfaction Survey Link"
  - "⚙️ Setup Survey Form Trigger"

**Build Process:** Run `node build.js` to regenerate ConsolidatedDashboard.gs

---

### Version 1.8.0 (2026-01-06) - Member Satisfaction Dashboard

**New Feature: Interactive Satisfaction Dashboard Modal**

Added a new modal popup dashboard for analyzing member satisfaction survey data, accessible via **Dashboard > 📊 Member Satisfaction**.

**Functions Added (Code.gs):**
- `showSatisfactionDashboard()` - Opens 900x750 modal dialog
- `getSatisfactionDashboardHtml()` - Returns HTML with 4-tab interface
- `getSatisfactionOverviewData()` - Calculates overview stats, NPS score, response rate
- `getSatisfactionResponseData()` - Fetches individual responses with filtering support
- `getSatisfactionSectionData()` - Fetches scores for all 11 survey sections
- `getSatisfactionAnalyticsData()` - Generates worksite/role analysis, steward impact, priorities

**Dashboard Tabs:**
1. **Overview** - 6 stat cards (responses, avg satisfaction, NPS, response rate, steward rating, leadership), circular gauge charts for key metrics, auto-generated insights
2. **Responses** - Searchable list with satisfaction level filtering (High ≥7, Medium 5-7, Low <5)
3. **By Section** - Bar chart showing all 11 survey sections ranked by score, section detail cards with progress bars
4. **Insights** - Key findings, satisfaction by worksite, by role, steward contact impact analysis, top member priorities

**Design:**
- Green gradient header theme (differentiates from purple Custom View)
- Mobile-responsive with 44px touch targets
- Client-side search and filtering for instant results
- Color-coded scores: green (7+), yellow (5-7), red (<5)
- Lazy-loads data per tab for performance

**Menu Location:** Dashboard > 📊 Member Satisfaction

---

### Version 1.7.0 (2026-01-03) - Dashboard Restructure & Bug Fixes

**Major Features:**

1. **Dashboard Restructure**
   - Removed unused columns G onwards - Dashboard now uses only columns A-F
   - Moved Steward Performance section to near bottom for better visual hierarchy
   - Added new **Top 30 Busiest Stewards** section showing stewards with most active cases
   - Compact Status Legend now fits in single row (A-F)
   - All numbers now formatted with comma separators (1,000)

2. **Overdue Cases Bug Fix**
   - Fixed Dashboard showing 0 for Overdue Cases
   - Root cause: Formula was looking for `<0` but Days to Deadline shows "Overdue" text
   - Changed formula to count cells containing "Overdue" text
   - Also fixed in hidden sheet Steward Workload, Location Analytics calculations

3. **Member Directory Enhancements**
   - Added collapsible column groups for Engagement Metrics (Q-T) and Member Interests (U-X)
   - Both groups are collapsed/hidden by default for cleaner view
   - Added conditional formatting: Has Open Grievance = "Yes" shows red background
   - Added comma formatting for numeric columns (Open Rate, Volunteer Hours)

4. **Tab Rename: Interactive → Custom View**
   - Renamed "🎯 Interactive" tab to "🎯 Custom View" for clarity
   - Updated SHEETS constant and all string references

5. **Code Cleanup**
   - Removed orphaned `setupInteractiveDashboardCalcSheet()` function (defined but never called)
   - Removed orphaned `setupEngagementCalcSheet()` function (used undefined constant)
   - Fixed hidden sheet number comments for accuracy

**Code Changes:**
- Code.gs: ~150 lines added for Dashboard restructure and Member Directory enhancements
- HiddenSheets.gs: Fixed 5 Overdue Cases formulas (changed `"<0"` to `"Overdue"`)
- HiddenSheets.gs: Removed 2 orphaned functions (~60 lines)
- Constants.gs: Changed `INTERACTIVE: '🎯 Interactive'` to `INTERACTIVE: '🎯 Custom View'`
- MobileQuickActions.gs: Updated "Interactive Dashboard" text references

---

### Version 1.6.0 (2026-01-03) - Desktop Search & Unified Seeding

**Major Features:**

1. **Desktop Search with Advanced Filtering**
   - Comprehensive search across Members and Grievances in one interface
   - Accessible via **Dashboard → 🔍 Search Members**
   - Tabbed interface: All, Members, Grievances
   - Advanced filters: Status (grievances), Location (all), Is Steward (members)
   - Searchable fields: Name, ID, Email, Job Title, Location, Issue Type, Steward
   - Click-to-navigate: Jump directly to any result row in the spreadsheet
   - Desktop optimized: 900x700 modal with responsive grid layout
   - Debounced search (300ms) for performance

2. **Dashboard STATUS Column Fix**
   - Fixed Dashboard metrics to use STATUS column for Win/Settled/Denied counts
   - Previously incorrectly referenced RESOLUTION column for some metrics
   - Status now includes both workflow states AND outcomes (single column design)
   - Ensures accurate grievance outcome tracking

3. **Unified Seed Function**
   - SEED_GRIEVANCES merged into SEED_MEMBERS function
   - Use `SEED_MEMBERS(count, grievancePercent)` to seed members with optional grievances
   - All seeded grievances are directly linked to members (no orphaned data)
   - Example: `SEED_MEMBERS(100)` seeds 100 members + ~30 grievances (30% default)
   - Example: `SEED_MEMBERS(100, 50)` seeds 100 members + ~50 grievances (50%)
   - Example: `SEED_MEMBERS(100, 0)` seeds members only, no grievances

4. **Automatic Timeline Column Grouping**
   - Grievance Log now automatically sets up collapsible column groups during creation
   - Step I columns (J-K) grouped together
   - Step II columns (L-O) grouped together
   - Step III columns (P-Q) grouped together
   - Click +/- controls to expand/collapse step details
   - Previously required manual setup via View menu

5. **Simplified Demo Menu**
   - Removed redundant "Seed Data" submenu
   - Single "🚀 Seed All Sample Data" option for seeding
   - Cleaner menu structure

6. **Member Directory Days to Deadline Fix**
   - Fixed Member Directory column AD not populating from Grievance Log
   - Root cause: MINIFS formula in _Grievance_Calc ignored "Overdue" text values
   - Solution: `syncGrievanceToMemberDirectory()` now calculates directly from Grievance Log
   - Properly handles both numeric deadlines and "Overdue" text
   - Shows minimum deadline when member has multiple open grievances

7. **SEED_SAMPLE_DATA Corrected**
   - Fixed to seed 1,000 members + 300 grievances (was incorrectly 50/25)
   - Uses merged approach: `SEED_MEMBERS(1000, 30)` for linked data
   - Automatically installs auto-sync trigger for live updates
   - Matches specification from DeveloperTools.gs

**New Functions:**
- `showDesktopSearch()` - Main desktop search dialog (~300 lines HTML/JS)
- `getDesktopSearchLocations()` - Get unique locations for filter dropdown
- `getDesktopSearchData(query, tab, filters)` - Backend search handler
- `navigateToSearchResult(type, id, row)` - Navigate to search result row

**Code Changes:**
- `Code.gs`: Updated `searchMembers()` to call `showDesktopSearch()`
- `ConsolidatedDashboard.gs`: Added desktop search functions
- `ConsolidatedDashboard.gs`: `createGrievanceLog()` now auto-creates column groups
- `ConsolidatedDashboard.gs`: Rewrote `syncGrievanceToMemberDirectory()` to calculate directly
- `ConsolidatedDashboard.gs`: Fixed `SEED_SAMPLE_DATA()` to seed 1000 members + 300 grievances
- `Constants.gs`: Updated `GRIEVANCE_STATUS` comment for clarity
- `HiddenSheets.gs`: Fixed Dashboard formulas to use STATUS column for outcome counts
- `DeveloperTools.gs`: Merged grievance seeding into SEED_MEMBERS function

**Desktop vs Mobile Search Comparison:**
| Aspect | Mobile | Desktop |
|--------|--------|---------|
| Search Fields | ID, Name, Email, Status | ID, Name, Email, Job Title, Location, Issue Type, Steward |
| Filters | Tab only | Status, Location, Is Steward |
| Result Limit | 20 | 50 |
| Navigation | No | Yes - click to jump to row |

---

### Version 1.5.1 (2026-01-03) - Mobile Web App for Phone Access

**Problem Solved:**
Google Sheets mobile app does not support Apps Script custom menus, making the dashboard and search features inaccessible on phones.

**Solution:**
Added standalone web app deployment that can be accessed via URL on any mobile browser.

**New File:**
- `WebApp.gs` (~510 lines) - Standalone web app with doGet() entry point

**Features:**
- **Dashboard Page**: Stats cards (Total, Active, Pending, Overdue grievances) + quick action buttons
- **Search Page**: Full-text search across members and grievances with tabbed filtering
- **Grievance List Page**: Filterable list by status (All, Open, Pending, Resolved)
- **Bottom Navigation**: Touch-friendly navigation bar
- **iOS Home Screen Support**: apple-mobile-web-app meta tags for app-like experience

**New Menu Item:**
- Dashboard → 📱 Get Mobile App URL - Shows deployment URL after web app is deployed

**Deployment Instructions:**
1. Extensions → Apps Script → Deploy → New deployment
2. Select "Web app"
3. Set access permissions
4. Copy URL and bookmark on mobile device

**Code Changes:**
- `WebApp.gs`: New file with doGet(), getWebAppDashboardHtml(), getWebAppSearchHtml(), getWebAppGrievanceListHtml(), getWebAppSearchResults(), getWebAppGrievanceList(), showWebAppUrl()
- `Code.gs`: Added menu item "📱 Get Mobile App URL" calling showWebAppUrl()

---

### Version 1.5.0 (2025-12-18) - Enhanced Analytics Dashboard

**Major Updates:**

- Enhanced 💼 Dashboard from 4 sections to 9 comprehensive analytics sections
- Added Operations Analytics-style metrics inspired by 509-dashboard
- All sections now use live auto-updating formulas (no manual refresh needed)

**New Dashboard Sections (5 added):**

| Section | Description |
|---------|-------------|
| TYPE ANALYSIS | Breakdown by issue category (Contract Violation, Discipline, Workload, Safety, Discrimination) with Total/Open/Resolved/Win Rate/Avg Days per category |
| LOCATION BREAKDOWN | Metrics per work location from Config - Members, Grievances, Open Cases, Win Rate |
| STEWARD PERFORMANCE | Total Stewards, Active w/Cases, Avg Cases/Steward, Vol Hours, Contacts This Month |
| MONTH-OVER-MONTH TRENDS | Filed/Closed/Won comparisons with 📈📉➡️ trend indicators and % change |
| STATUS LEGEND | Moved from section 5 to section 9 |

**Formula Types Used:**

- `COUNTIF` / `COUNTIFS` - Count by criteria
- `AVERAGEIFS` - Average by criteria
- `IFERROR` / `IF` - Error handling
- `TEXT` - Percentage formatting
- Date math with `DATE()`, `YEAR()`, `MONTH()`, `TODAY()`

**Live Data Flow:**

```
Config (Office Locations) ──┐
                            ├──► 💼 Dashboard (auto-updates)
Member Directory ───────────┤
                            │
Grievance Log ──────────────┘
```

**Code Changes:**

- Code.gs: Enhanced `createDashboard()` from ~170 lines to ~320 lines
- Code.gs: Fixed `rebuildDashboard()` to actually recreate dashboard sheets
- AIR.md: Updated documentation with 9-section dashboard architecture

---

### Version 1.4.5 (2025-12-18) - Auto-Sort & Seed Improvements

**New Feature: Auto-Sort Grievance Log by Status**
- Grievance Log now automatically sorts when edited
- Primary sort: Status priority (active cases first, resolved cases last)
- Secondary sort: Days to deadline (most urgent first within each status)

**Status Priority Order:**
| Priority | Status | Type |
|----------|--------|------|
| 1 | Open | Active |
| 2 | Pending Info | Active |
| 3 | In Arbitration | Active |
| 4 | Appealed | Active |
| 5 | Settled | Resolved |
| 6 | Won | Resolved |
| 7 | Denied | Resolved |
| 8 | Withdrawn | Resolved |
| 9 | Closed | Resolved |

**Seed Data Improvements:**
- Expanded name pools from 20 to 120 names each (14,400+ unique combinations)
- Significantly reduced repetition of names in seeded member and grievance data

**Code Changes:**
- `Constants.gs`: Added `GRIEVANCE_STATUS_PRIORITY` constant
- `HiddenSheets.gs`: Added `sortGrievanceLogByStatus()` function
- `HiddenSheets.gs`: Hooked sort into `onEditAutoSync()` trigger
- `ConsolidatedDashboard.gs`: Added duplicate `sortGrievanceLogByStatus()` function and trigger hook
- `DeveloperTools.gs`: Expanded `firstNames` and `lastNames` arrays (20 → 120 each)
- `ConsolidatedDashboard.gs`: Expanded name arrays in `SEED_MEMBERS()` function

---

### Version 1.4.4 (2025-12-18) - Grievance Log Member Lookup Fix

**Grievance Log now auto-populates member data directly from Member Directory:**

| Column | Header | Auto-Synced From |
|--------|--------|------------------|
| C | First Name | Member Directory (by Member ID) |
| D | Last Name | Member Directory (by Member ID) |
| X | Member Email | Member Directory (by Member ID) |
| Y | Unit | Member Directory (by Member ID) |
| Z | Work Location | Member Directory (by Member ID) |
| AA | Steward | Member Directory (by Member ID) |

**Member ID Entry:**
- Column B (Member ID) allows free text entry (no dropdown restriction)
- Member IDs should match Member Directory entries for auto-lookup to work
- Invalid/mismatched Member IDs will result in empty lookup fields (C-D, X-AA)

**Code Changes:**
- `HiddenSheets.gs`: Rewrote `syncGrievanceFormulasToLog()` to lookup member data directly from Member Directory instead of using hidden sheet formulas
- Bypassed the ARRAYFORMULA/FILTER issue that caused empty lookups
- Member lookup now uses `MEMBER_COLS` constants for reliable column access

---

### Version 1.4.3 (2025-12-18) - Member Directory Auto-Sync

**Member Directory Columns AB-AD now auto-sync from Grievance Log:**

| Column | Header | Auto-Synced Value |
|--------|--------|-------------------|
| AB | Has Open Grievance? | "Yes" or "No" based on active grievances |
| AC | Grievance Status | Status from most recent grievance |
| AD | Days to Deadline | Countdown number or "Overdue" |

**Code Changes:**
- `HiddenSheets.gs`: Changed `_Grievance_Calc` hidden sheet to pull `DAYS_TO_DEADLINE` instead of `NEXT_ACTION_DUE`
- `Constants.gs`: Updated header from "Next Deadline" to "Days to Deadline"
- Auto-sync trigger updates Member Directory when Grievance Log is edited

---

### Version 1.4.2 (2025-12-18) - Date Formatting & Overdue Display

**Formatting Changes:**
- Date format changed from `yyyy-mm-dd` to `dd-mm-yyyy` throughout
- Days Open (S) and Days to Deadline (U) now display as whole numbers (no decimals)
- Days to Deadline shows "Overdue" for past-due cases instead of negative numbers

**Display Values for Days to Deadline:**
| Value | Meaning |
|-------|---------|
| `18` | 18 days remaining |
| `0` | Due today |
| `Overdue` | Past deadline |
| *(blank)* | Case is closed |

**Code Changes:**
- `Code.gs`: Added `setNumberFormat('0')` for Days Open and Days to Deadline in `createGrievanceLog()`
- `Code.gs`: Changed date format to `dd-mm-yyyy` in `createGrievanceLog()`
- `HiddenSheets.gs`: Changed date format to `dd-mm-yyyy` in all sync functions
- `HiddenSheets.gs`: Days to Deadline now returns "Overdue" when `days < 0`
- `Code.gs`: Updated setup success message to confirm auto-sync trigger installation

---

### Version 1.4.1 (2025-12-16) - Days to Deadline Fix

**Bug Fix:**
- Fixed "Days to Deadline" and "Days Open" showing duplicate/incorrect values for all grievances
- Root cause: ARRAYFORMULA with FILTER-based row index in hidden sheet didn't expand correctly
- Solution: Calculate Days Open, Next Action Due, Days to Deadline, and all deadline dates directly in JavaScript within `syncGrievanceFormulasToLog()` function

**Code Changes:**
- `HiddenSheets.gs`: Rewrote metrics calculation in `syncGrievanceFormulasToLog()` (~60 lines added)
  - Now calculates Filing Deadline, Step I/II/III Due dates from source dates
  - Days Open = (Date Closed or Today) - Date Filed
  - Next Action Due = Based on Current Step status
  - Days to Deadline = Next Action Due - Today
  - All values now calculated per-row from actual grievance data

---

### Version 1.4.0 (2025-12-16) - Dashboard Views Added

**Major Updates:**

- Re-added dashboard sheets from original 509dashboard project
- Created unified 💼 Dashboard (merged Executive Dashboard + Dashboard themes)
- Added 🎯 Custom View with customizable metric selection
- Added `_Dashboard_Calc` hidden sheet with 15 self-healing metric formulas

**New Sheets (2):**

- `💼 Dashboard` - Executive-style metrics view with:
  - QUICK STATS section (green Union theme)
  - MEMBER METRICS section (blue theme)
  - GRIEVANCE METRICS section (orange theme)
  - TIMELINE & PERFORMANCE section (purple theme)
  - Real-time formulas linked to Member Directory and Grievance Log

- `🎯 Interactive` - Customizable dashboard with:
  - Dropdown metric selection (8 available metrics)
  - Time range filtering
  - Theme selection
  - Live-updating values

**New Hidden Sheet:**

- `_Dashboard_Calc` - 15 key metrics with self-healing formulas:
  - Total Members, Active Stewards
  - Total/Open/Pending/Settled/Won/Denied/Withdrawn Grievances
  - Win Rate %, Avg Days to Resolution
  - Overdue Cases, Due This Week
  - Filed This Month, Closed This Month

**Code Changes:**

- Constants.gs: Added DASHBOARD, INTERACTIVE, DASHBOARD_CALC to SHEETS
- Code.gs: Added `createDashboard()`, `createInteractiveDashboard()` (~300 lines)
- HiddenSheets.gs: Added `setupDashboardCalcSheet()` (~70 lines)
- Updated CREATE_509_DASHBOARD, DIAGNOSE_SETUP, REPAIR_DASHBOARD for 5 sheets

---

### Version 1.1.0 (2025-12-14) - Hidden Sheet Architecture

Added self-healing hidden formula system:

- HiddenSheets.gs: Full implementation of 6 hidden calculation sheets
- Auto-sync onEdit trigger for automatic cross-sheet data population
- Self-healing repair functions
- Manual sync options in Administrator menu

### Version 1.0.0 (2025-12-14) - Fresh Start

Complete rebuild from AIR.md specification.

- Removed 77 legacy .gs files (~118,000 lines)
- Created 3 clean files (~1,500 lines)
- Constants.gs: Column mappings and helper functions
- Code.gs: Setup, menus, and sheet creation
- DeveloperTools.gs: Demo data management (DELETE BEFORE PRODUCTION)
- 98.8% code reduction while maintaining core functionality
