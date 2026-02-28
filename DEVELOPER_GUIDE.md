# Dashboard Developer Guide

This guide provides an overview of the codebase architecture, development workflow, and best practices for contributing to the Union Steward Dashboard.

## Table of Contents

1. [Project Overview](#project-overview)
2. [Architecture](#architecture)
3. [Module Structure](#module-structure)
4. [Security](#security)
5. [Data Access Layer](#data-access-layer)
6. [Build System](#build-system)
7. [Testing](#testing)
8. [Code Style Guidelines](#code-style-guidelines)
9. [Common Patterns](#common-patterns)
10. [Debugging](#debugging)

---

## Project Overview

The Dashboard is a Google Apps Script (GAS) application for managing union steward activities, grievances, and member information within Google Sheets.

**Key Technologies:**
- Google Apps Script (V8 Runtime with ES2020 support)
- Google Sheets API
- Node.js build system
- ESLint for code quality (v9.x flat config)

**Version:** 4.13.0

---

## Architecture

### High-Level Structure

```
DDS-Dashboard/
├── src/                    # 37 source files (.gs) + 7 HTML
├── test/                   # Jest unit tests (1300+ tests)
├── dist/                   # Build output (auto-generated)
├── setup-instructions/     # Optional feature setup guides
├── .github/workflows/      # CI/CD configuration
├── build.js               # Build script
├── jest.config.js         # Test configuration
├── eslint.config.js       # Linting configuration (v9 flat config)
├── package.json           # Dependencies and scripts
└── DEVELOPER_GUIDE.md     # This file
```

### Modular Design (v4.6.0)

The codebase follows a layered architecture with 37 modules organized by numbered prefix:

```
┌─────────────────────────────────────────────────────────────┐
│                      10_Main.gs                              │
│               (Entry Point / onOpen / onEdit)                │
└─────────────────────────────────────────────────────────────┘
                              │
        ┌─────────────────────┼──────────────────────┐
        ▼                     ▼                      ▼
┌───────────────┐   ┌────────────────┐   ┌────────────────────┐
│  UI Layer     │   │  Data Layer    │   │  Services/Features  │
│  03/04a-e     │   │  02/08a-d/10*  │   │  05/06/09/11-14    │
└───────────────┘   └────────────────┘   └────────────────────┘
        │                     │                      │
        └─────────────────────┼──────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│               00_DataAccess.gs (Data Access Layer)           │
│            Cached sheet access, batch operations             │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                00_Security.gs (Security Layer)               │
│     XSS prevention, access control, PII masking, validation  │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
                    ┌───────────────┐
                    │  01_Core.gs   │
                    │ (Error + Config)│
                    └───────────────┘
```

---

## Module Structure

### Foundation Modules (Load First)

| File | Purpose | Key Functions |
|------|---------|---------------|
| `00_Security.gs` | Security utilities, XSS prevention, access control | `escapeHtml`, `checkWebAppAuthorization`, `maskEmail`, `validateWebAppRequest` |
| `00_DataAccess.gs` | Data Access Layer, time constants, cached sheet access | `DataAccess.getSheet`, `DataAccess.getMemberById`, `TIME_CONSTANTS` |
| `01_Core.gs` | Error handling + Constants | `handleError`, `SHEETS`, `MEMBER_COLS`, `GRIEVANCE_COLS` |

### Data Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| `02_DataManagers.gs` | Member + Grievance CRUD | `addMember`, `updateMember`, `createGrievance` |
| `08a_SheetSetup.gs` | Main setup, utility functions, data validation, multi-select | `getOrCreateSheet`, `setupDataValidation`, `onEditMultiSelect` |
| `08b_SearchAndCharts.gs` | Search functions, chart generation | `desktopSearch`, `buildChart` |
| `08c_FormsAndNotifications.gs` | Form handling, notifications, deadline alerts | `sendDeadlineAlerts`, `processFormSubmission` |
| `08d_AuditAndFormulas.gs` | Audit log, formula sync, hidden calc sheets | `logAuditEntry`, `setupCalcFormulasSheet` |

### UI Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| `03_UIComponents.gs` | Menu, Theme, Mobile, QuickActions, Search | `createDashboardMenu`, `applyTheme`, `showQuickActions` |
| `04a_UIMenus.gs` | Menu creation, visual control panel, navigation, dialogs | `buildMenuStructure`, `showNavigationSidebar` |
| `04b_AccessibilityFeatures.gs` | Comfort view, focus mode, import/export | `setupComfortView`, `showComfortViewPanel` |
| `04c_InteractiveDashboard.gs` | Interactive dashboard, mobile views, data retrieval | `showInteractiveDashboard`, `getMobileViewData` |
| `04d_ExecutiveDashboard.gs` | Executive dashboard, steward dashboard, alerts | `showExecutiveDashboard`, `showStewardAlerts` |
| `04e_PublicDashboard.gs` | Public member dashboard, unified data endpoints | `getUnifiedDashboardData`, `buildMemberDashboard` |

### Service Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| `05_Integrations.gs` | Drive, Calendar, WebApp, Constant Contact integration | `doGet`, `syncToCalendar`, `setupDriveFolder`, `createMeetingDocs`, `syncConstantContactEngagement` |
| `06_Maintenance.gs` | Diagnostics, Cache, Undo | `DIAGNOSE_SETUP`, `getCachedData`, `undoLastAction` |
| `09_Dashboards.gs` | Satisfaction, Sync, Public dashboards | `getSatisfactionData`, `onEditAutoSync`, `buildPublicPortal` |

### Business Logic & Entry Point

| File | Purpose | Key Functions |
|------|---------|---------------|
| `10_Main.gs` | Entry point + Triggers | `onOpen`, `onEdit` (consolidated dispatcher), `dailyTrigger` |
| `10a_SheetCreation.gs` | Config, Member Directory, Grievance Log sheet creation | `createConfigSheet`, `createMemberDirectorySheet` |
| `10b_SurveyDocSheets.gs` | Satisfaction, Feedback, FAQ, Getting Started sheets | `createSatisfactionSheet`, `createFAQSheet` |
| `10c_FormHandlers.gs` | Menu handlers, form submissions, flagged reviews | `onGrievanceFormSubmit`, `handleFlaggedReview` |
| `10d_SyncAndMaintenance.gs` | Formatting, testing, Drive, Calendar, Email, sync | `syncAllData`, `applyGlobalFormatting` |

### Feature Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| `11_CommandHub.gs` | Command center + Secure dashboard | `showCommandCenter`, `buildSecureDashboard` |
| `12_Features.gs` | Dynamic Engine, Looker Studio, Checklist | `handleChecklistEdit`, `buildSafeQuery`, `setupLookerIntegration` |
| `13_MemberSelfService.gs` | Member self-service portal with PIN authentication | `generateMemberPIN`, `authenticateMember`, `showMemberSelfService` |
| `14_MeetingCheckIn.gs` | Meeting check-in system with email + PIN auth | `showSetupMeetingDialog`, `createMeeting`, `showMeetingCheckInDialog` |

### Event Bus & Analytics Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| `15_EventBus.gs` | Pub/sub event system | `EventBus.emit`, `EventBus.on`, `EventBus.off` |
| `16_DashboardEnhancements.gs` | Date ranges, chart export, drill-down | `showEnhancedDashboard`, `exportChart` |
| `17_CorrelationEngine.gs` | Cross-dimensional correlation analysis | `calculateCorrelation`, `getInsightStrings` |

### SPA Web Dashboard Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| `19_WebDashAuth.gs` | Google SSO + magic link auth | `initWebDashboardAuth`, `verifyMagicLink` |
| `20_WebDashConfigReader.gs` | Column-based Config tab reader | `getConfigValue`, `CONFIG_COLS` |
| `21_WebDashDataService.gs` | Unified data service for SPA | `getSPADashboardData`, `getMemberViewData` |
| `22_WebDashApp.gs` | SPA entry point and routing | `doGetWebDashboard`, `PAGE_DATA` |
| `23_PortalSheets.gs` | Hidden sheet management for SPA | `getPortalSheetData`, `updatePortalSheet` |
| `24_WeeklyQuestions.gs` | Weekly check-in questions | `getWeeklyQuestions`, `submitWeeklyResponse` |
| `25_WorkloadService.gs` | SPA-integrated workload (IIFE: WorkloadService) | `getWorkloadFormData`, `submitWorkloadData` |

### HTML Templates

| File | Purpose |
|------|---------|
| `MultiSelectDialog.html` | Multi-select dialog HTML template |
| `index.html` | SPA entry point |
| `styles.html` | SPA shared styles |
| `steward_view.html` | Steward dashboard SPA view |
| `member_view.html` | Member dashboard SPA view |
| `portal_sheets.html` | Portal sheets management UI |
| `weekly_questions.html` | Weekly questions UI |

### Development Module (Remove in Production)

| File | Purpose | Key Functions |
|------|---------|---------------|
| `07_DevTools.gs` | Test data generation (delete before go-live) | `SEED_TEST_DATA`, `NUKE_SEEDED_DATA` |

---

## Security

### Overview

The `00_Security.gs` module provides centralized security functions:

### XSS Prevention

```javascript
// Server-side HTML escaping
var safeName = escapeHtml(userData.name);

// For client-side templates, include the security script
var html = getClientSecurityScript() + '<body>...</body>';

// Pre-sanitize data arrays before passing to client
var safeData = sanitizeDataForClient(members, ['firstName', 'lastName', 'email']);
```

### Access Control

```javascript
// In doGet() - validate request and check authorization
var validation = validateWebAppRequest(e);
if (!validation.isValid) {
  return getAccessDeniedPage('Invalid request');
}

var authResult = checkWebAppAuthorization('steward');
if (!authResult.isAuthorized) {
  return getAccessDeniedPage(authResult.message);
}
```

### Security Event Alerting

The `00_Security.gs` module includes a security event alerting system that detects and notifies on suspicious activity. Hooks are integrated at key entry points:

- **`05_Integrations.gs`** — web app access monitoring
- **`10_Main.gs`** — edit trigger threat detection
- **`13_MemberSelfService.gs`** — failed authentication attempt tracking

### Member ID Generation

Member IDs use a **name-based format** generated by `generateNameBasedId()`:

- **Format:** `M` + first 2 letters of first name + first 2 letters of last name + 3 random digits (e.g., `MJASM472` for Jane Smith)
- **Collision handling:** Up to 100 retries with new digits, then falls back to 8 hex UUID characters
- **Import dedup:** Happens once client-side before batch processing for performance
- IDs are generated via Strategic Ops > ID Engines > Generate Missing IDs

### PII Masking for Logs

```javascript
// Use secure logging to mask PII
secureLog('processGrievance', 'Processing grievance', {
  email: member.email,     // Will be masked as "j***@example.com"
  phone: member.phone,     // Will be masked as "***-***-1234"
  firstName: member.firstName  // Will be masked as "J."
});
```

### Formula Injection Prevention

```javascript
// Use safe formula builders
var formula = buildSafeQuery(sheetName, 'SELECT A, B WHERE C > 0', 1);

// Escape sheet names for formula use
var safeSheet = safeSheetNameForFormula(userProvidedName);
```

---

## Data Access Layer

### Overview

The `00_DataAccess.gs` module provides a centralized data access layer with caching:

### Basic Usage

```javascript
// Get a cached sheet reference
var sheet = DataAccess.getSheet(SHEETS.MEMBER_DIR);

// Get all data (cached)
var members = DataAccess.getAllMembers();

// Get specific member
var member = DataAccess.getMemberById('MEM-001');

// Find with predicate
var overdueGrievances = DataAccess.findAllRows(SHEETS.GRIEVANCE_LOG, function(row) {
  return row[GRIEVANCE_COLUMNS.DAYS_TO_DEADLINE] < 0;
});
```

### Time Constants

```javascript
// Use TIME_CONSTANTS instead of magic numbers
var filingDeadline = calculateDeadline(incidentDate, TIME_CONSTANTS.DEADLINE_DAYS.FILING);

// Get urgency level
var urgency = getDeadlineUrgency(daysToDeadline);  // 'critical', 'warning', 'normal', 'overdue'
```

### Cache Invalidation

```javascript
// After structural changes (adding/removing sheets)
DataAccess.invalidateCache();
```

---

## Build System

### Commands

```bash
# Install dependencies
npm install

# Run linting only
npm run lint

# Run linting with auto-fix
npm run lint:fix

# Build consolidated file
npm run build

# Lint + Build + Jest tests
npm run test

# Run Jest unit tests only
npm run test:unit

# Build with linting
node build.js --lint

# Clean dist directory
npm run clean

# Deploy to Google Apps Script
npm run deploy
```

### Build Process

1. **Lint** (optional): ESLint checks all `.gs` files
2. **Concatenate**: Files combined in `BUILD_ORDER` sequence
3. **Embed HTML**: HTML templates converted to functions
4. **Output**: Individual `.gs` + `.html` files copied to `dist/`

### Build Order

Files must be concatenated in dependency order (37 files):

```javascript
const BUILD_ORDER = [
  '00_Security.gs',               // Security utilities, XSS prevention, access control
  '00_DataAccess.gs',             // Data Access Layer, time constants
  '01_Core.gs',                   // Error handling + Constants (SHEETS, MEMBER_COLS, etc.)
  '02_DataManagers.gs',           // Member + Grievance CRUD managers
  '03_UIComponents.gs',           // Menu, Theme, Mobile, QuickActions, Search
  '04a_UIMenus.gs',               // Menu creation, visual control panel, dialogs
  '04b_AccessibilityFeatures.gs', // Comfort view, focus mode, import/export
  '04c_InteractiveDashboard.gs',  // Interactive dashboard, mobile views
  '04d_ExecutiveDashboard.gs',    // Executive dashboard, steward dashboard, alerts
  '04e_PublicDashboard.gs',       // Public member dashboard, unified data endpoints
  '05_Integrations.gs',           // Drive, Calendar, WebApp integration
  '06_Maintenance.gs',            // Diagnostics, Cache, Undo/Redo
  '07_DevTools.gs',               // Dev tools + Test framework (remove before prod)
  '08a_SheetSetup.gs',            // Main setup, utility functions, data validation
  '08b_SearchAndCharts.gs',       // Search functions, chart generation
  '08c_FormsAndNotifications.gs', // Form handling, notifications, deadline alerts
  '08d_AuditAndFormulas.gs',      // Audit log, formula sync, hidden calc sheets
  '09_Dashboards.gs',             // Satisfaction, Sync, Public dashboards
  '10a_SheetCreation.gs',         // Config, Member Directory, Grievance Log creation
  '10b_SurveyDocSheets.gs',       // Satisfaction, Feedback, FAQ, Getting Started sheets
  '10c_FormHandlers.gs',          // Menu handlers, form submissions, flagged reviews
  '10d_SyncAndMaintenance.gs',    // Formatting, testing, Drive, Calendar, sync
  '10_Main.gs',                   // Main entry point, onOpen, onEdit, triggers
  '11_CommandHub.gs',             // Command center + Secure member dashboard
  '12_Features.gs',               // Checklist, Dynamic Engine, Looker Studio
  '13_MemberSelfService.gs',      // Member self-service portal with PIN auth
  '14_MeetingCheckIn.gs',         // Meeting check-in system with email + PIN auth
  '15_EventBus.gs',               // Pub/sub event system
  '16_DashboardEnhancements.gs',  // Date ranges, chart export, drill-down
  '17_CorrelationEngine.gs',      // Cross-dimensional correlation analysis
  '19_WebDashAuth.gs',            // Google SSO + magic link auth
  '20_WebDashConfigReader.gs',    // Column-based Config tab reader
  '21_WebDashDataService.gs',     // Unified data service for SPA
  '22_WebDashApp.gs',             // SPA entry point and routing
  '23_PortalSheets.gs',           // Hidden sheet management for SPA
  '24_WeeklyQuestions.gs',        // Weekly check-in questions
  '25_WorkloadService.gs'         // SPA-integrated workload (IIFE: WorkloadService)
];
```

---

## Testing

### Jest Test Suite (Primary)

The project uses **Jest v29.7.0** as its primary test framework, with 1300+ tests across 21+ test suites. Tests run in Node.js using a GAS mock infrastructure that simulates the Google Apps Script environment.

#### Running Tests

```bash
# Run Jest unit tests only
npm run test:unit

# Run full pipeline: lint + build + Jest tests
npm test
```

#### Test Directory Structure

```
test/
├── gas-mock.js                    # GAS global mocks (SpreadsheetApp, Logger, etc.)
├── load-source.js                 # Loads .gs source files into Node.js global scope
├── modules.test.js                # Cross-module integration tests
├── 00_DataAccess.test.js          # Data Access Layer tests
├── 00_Security.test.js            # Security module tests
├── 01_Core.test.js                # Core constants and error handling tests
├── 02_DataManagers.test.js        # Member and grievance data manager tests
├── 03_UIComponents.test.js        # UI component tests
├── 04_UIService.test.js           # UI service tests
├── 04e_PublicDashboard.test.js    # Public dashboard tests
├── 05_Integrations.test.js        # Integrations tests
├── 06_Maintenance.test.js         # Maintenance and diagnostics tests
├── 07_DevTools.test.js            # Dev tools tests
├── 08_SheetUtils.test.js          # Sheet utility tests
├── 09_Dashboards.test.js          # Dashboard tests
├── 10_Code.test.js                # Code and sheet creation tests
├── 10_Main.test.js                # Main entry point tests
├── 11_CommandHub.test.js          # Command hub tests
├── 12_Features.test.js            # Features and checklist tests
├── 13_MemberSelfService.test.js   # Member self-service portal tests
└── 14_MeetingCheckIn.test.js      # Meeting check-in system tests
```

#### Writing Tests

```javascript
// test/example.test.js
require('./gas-mock');                     // Set up GAS environment mocks
const { loadSource } = require('./load-source');

beforeAll(() => {
  loadSource('01_Core.gs');                // Load source file into global scope
});

describe('myFunction', () => {
  it('should return expected value', () => {
    var result = global.myFunction(42);
    expect(result).toBe(42);
  });
});
```

The `gas-mock.js` file provides mock implementations for GAS globals (`SpreadsheetApp`, `Logger`, `Utilities`, `DriveApp`, etc.) so that source files can be loaded and tested outside of the Apps Script runtime.

The `load-source.js` helper rewrites top-level `var` and `function` declarations to `global.*` assignments so they become accessible in Jest's sandbox.

### Custom GAS Test Framework (Legacy/In-Editor)

The `07_DevTools.gs` module also includes a custom in-editor test framework for running tests directly within the Apps Script environment:

```javascript
// Run all tests in Apps Script editor
runAllTests();

// Show test dashboard UI
showTestDashboard();
```

---

## Code Style Guidelines

### General Rules

1. **Prefer `var` for GAS compatibility** - V8 runtime supports `let`/`const` but `var` is used by convention
2. **Arrow functions are supported** - but `function` keyword is used by project convention
3. **JSDoc comments** - Document all public functions
4. **Semicolons required** - End statements with `;`

### Naming Conventions

```javascript
// Functions: camelCase
function getUserById(userId) { }

// Menu-callable functions: UPPER_SNAKE_CASE
function CREATE_DASHBOARD() { }

// Private functions: camelCase with trailing underscore
function helperFunction_() { }

// Constants: UPPER_SNAKE_CASE
var MAX_RETRIES = 3;

// Configuration objects: UPPER_SNAKE_CASE
var CACHE_CONFIG = { TTL: 300 };
```

### Function Documentation

```javascript
/**
 * Brief description of what the function does
 * @param {string} param1 - Description of param1
 * @param {number} [param2=10] - Optional param with default
 * @returns {Object} Description of return value
 * @throws {Error} When something goes wrong
 */
function exampleFunction(param1, param2) {
  param2 = param2 || 10;
  // implementation
}
```

---

## Common Patterns

### Getting Sheet Data

```javascript
function getSheetData(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
}
```

### Using the Cache Layer

```javascript
// With caching
var data = getCachedData(CACHE_KEYS.ALL_MEMBERS, function() {
  return loadMembersFromSheet();
}, 600);

// Invalidate when data changes
invalidateCache(CACHE_KEYS.ALL_MEMBERS);
```

### Recording Undo Actions

```javascript
// Before making changes
var beforeState = { row: rowNum, col: colNum, value: oldValue };

// Make the change
sheet.getRange(rowNum, colNum).setValue(newValue);

// Record for undo
recordAction('EDIT_CELL', 'Edited cell', beforeState, {
  row: rowNum,
  col: colNum,
  value: newValue
});
```

### Error Handling

```javascript
function safeOperation() {
  try {
    // risky operation
    return doSomething();
  } catch (e) {
    Logger.log('Error in safeOperation: ' + e.message);
    SpreadsheetApp.getActiveSpreadsheet()
      .toast('Operation failed: ' + e.message, 'Error', 5);
    return null;
  }
}
```

### External API Integration (Constant Contact)

The CC integration in `05_Integrations.gs` demonstrates the pattern for external REST API calls:

```javascript
// 1. Get a valid token (auto-refreshes if expired)
var token = getConstantContactToken_();

// 2. Make authenticated API calls via the helper
var data = ccApiGet_('/contacts', { limit: 500, include: 'email_address' });

// 3. The helper handles: Bearer auth, 401 retry with refresh, 429 rate limit retry
```

**Key patterns used:**
- **Token storage**: `PropertiesService.getScriptProperties()` for encrypted-at-rest credential storage
- **Token refresh**: Automatic refresh when `CC_TOKEN_EXPIRY` is past, using the stored refresh token
- **Rate limiting**: `Utilities.sleep(CC_CONFIG.RATE_LIMIT_DELAY_MS)` between paginated API calls
- **Error handling**: `muteHttpExceptions: true` + manual response code checking for graceful degradation

**Testing external APIs:**
```javascript
// In beforeEach, clear mock and set up valid token
UrlFetchApp.fetch.mockClear();
UrlFetchApp.fetch.mockReturnValueOnce({
  getResponseCode: jest.fn(() => 200),
  getContentText: jest.fn(() => JSON.stringify({ contacts: [] }))
});
```

The `UrlFetchApp` mock is defined in `test/gas-mock.js` and returns `200 + '{}'` by default. Override with `mockReturnValueOnce()` for specific test scenarios.

---

## Debugging

### Using Logger

```javascript
Logger.log('Debug message');
Logger.log('Value: ' + JSON.stringify(obj));
```

### View Logs

1. In Apps Script editor: View > Logs
2. Or use `console.log()` (visible in browser dev tools when using sidebars)

### Diagnostic Tools

```javascript
// Run full diagnostics
DIAGNOSE_SETUP();

// Show diagnostic dialog
showDiagnosticsDialog();

// Check cache status
showCacheStatusDashboard();
```

### Common Issues

| Issue | Solution |
|-------|----------|
| "Sheet not found" | Check `SHEETS` constant in `01_Core.gs` |
| Stale data | Call `invalidateAllCaches()` |
| Trigger not firing | Run `DIAGNOSE_SETUP()` to check triggers |
| UI not updating | Call `SpreadsheetApp.flush()` |

---

## CI/CD Pipeline

The GitHub Actions workflow (`.github/workflows/build.yml`) runs on every push:

1. **Install dependencies** - `npm ci`
2. **Lint** - `npm run lint`
3. **Build** - `npm run build`
4. **Verify** - Check build output exists
5. **Report** - Generate build summary
6. **Archive** - Upload build artifact

### Manual Deployment

```bash
# Full deploy with linting
npm run deploy

# Or step by step:
npm run lint
npm run build
clasp push
```

---

## Contributing

1. Create a feature branch from `main`
2. Make changes in `src/` directory
3. Run `npm run lint` to check for issues
4. Run `npm run build` to test build
5. Commit with descriptive messages
6. Push and create PR

### Commit Message Format

```
type: brief description

- Detailed change 1
- Detailed change 2

Closes #issue-number
```

Types: `feat`, `fix`, `refactor`, `docs`, `test`, `chore`
