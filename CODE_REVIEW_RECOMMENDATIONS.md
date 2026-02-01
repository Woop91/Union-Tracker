# Code Review Recommendations

**Repository:** Union Steward Dashboard (v4.4.1)
**Review Date:** February 2026
**Reviewer:** Claude Code

---

## Executive Summary

This comprehensive code review analyzed 13 Google Apps Script modules totaling 46,308 lines of code. While the project demonstrates solid functionality and has made progress in modularization, there are significant opportunities for improvement in security, code quality, architecture, and testing.

### Overall Health Scores

| Category | Score | Status |
|----------|-------|--------|
| Security | 4/10 | Critical issues requiring immediate attention |
| Code Quality | 5/10 | Moderate issues affecting maintainability |
| Architecture | 4/10 | Significant coupling and separation concerns |
| Testing | 3/10 | Minimal coverage (~3%) |
| Documentation | 6/10 | Good but partially outdated |
| Technical Debt | 8/10 | Well-managed (only 1 TODO found) |

---

## Table of Contents

1. [Critical Security Issues](#1-critical-security-issues)
2. [Code Quality Issues](#2-code-quality-issues)
3. [Architecture Recommendations](#3-architecture-recommendations)
4. [Performance Optimizations](#4-performance-optimizations)
5. [Testing Improvements](#5-testing-improvements)
6. [Documentation Updates](#6-documentation-updates)
7. [Build System Issues](#7-build-system-issues)
8. [Prioritized Action Items](#8-prioritized-action-items)

---

## 1. Critical Security Issues

### 1.1 Cross-Site Scripting (XSS) Vulnerabilities

**Severity: CRITICAL**

Multiple instances of unsanitized data being directly injected into `innerHTML`:

| File | Lines | Issue |
|------|-------|-------|
| `09_Dashboards.gs` | 316-331 | User data (worksite, role, shift) directly concatenated |
| `05_Integrations.gs` | 1499-1501 | Member/grievance fields injected without sanitization |
| `10_Main.gs` | 1788, 1791 | Error messages directly concatenated |
| `04_UIService.gs` | 2457, 2478 | Member data rendered without escaping |

**Recommendation:**
```javascript
// BAD - Current code
c.innerHTML = "<div>" + r.worksite + "</div>";

// GOOD - Use textContent or proper escaping
function escapeHtml(text) {
  var div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}
c.innerHTML = "<div>" + escapeHtml(r.worksite) + "</div>";
```

### 1.2 Formula Injection Vulnerabilities

**Severity: CRITICAL**

Configuration values are directly interpolated into Google Sheets formulas without validation:

| File | Lines | Issue |
|------|-------|-------|
| `12_Features.gs` | 1647, 1651-1652 | Sheet names in QUERY formulas |
| `08_SheetUtils.gs` | 3377, 3720, 4080, 4147 | Template literals with unsanitized sheet names |

**Recommendation:**
```javascript
// Create a formula-safe escaping function
function escapeForFormula(input) {
  if (!input) return '';
  return String(input)
    .replace(/'/g, "''")
    .replace(/"/g, '""')
    .replace(/\\/g, '\\\\');
}
```

### 1.3 Missing Access Control in Web App

**Severity: CRITICAL**

**Location:** `05_Integrations.gs` - `doGet()` function (lines 1114-1141)

The web app has no authentication or authorization mechanism:
- No verification that users have permission to view data
- The `mode` parameter determines PII access without validation
- No role-based access control (RBAC)

**Recommendation:**
```javascript
function doGet(e) {
  // Add authorization check
  var user = Session.getEffectiveUser().getEmail();
  if (!isAuthorizedUser(user)) {
    return HtmlService.createHtmlOutput('<h1>Access Denied</h1>');
  }

  // Validate mode parameter against allowlist
  var allowedModes = ['steward', 'member', 'dashboard'];
  var mode = e && e.parameter && e.parameter.mode;
  if (mode && allowedModes.indexOf(mode) === -1) {
    return HtmlService.createHtmlOutput('<h1>Invalid Request</h1>');
  }
  // ... rest of logic
}
```

### 1.4 Data Exposure Through Logging

**Severity: MEDIUM**

Personal identifiable information (PII) is being logged:

| File | Line | Data Exposed |
|------|------|--------------|
| `05_Integrations.gs` | 897 | Email addresses |
| `09_Dashboards.gs` | 1823, 1886, 1889 | Member IDs, Grievance IDs |
| `01_Core.gs` | 176 | Admin email addresses |

**Recommendation:** Remove PII from logs or implement log sanitization.

### 1.5 Weak Sanitization Function

**Location:** `01_Core.gs` lines 216-220

The `sanitizeForQuery()` function only escapes quotes and backslashes, which is insufficient for formula injection and other contexts.

**Recommendation:** Use context-specific sanitization functions or a dedicated security library.

---

## 2. Code Quality Issues

### 2.1 Inconsistent Constant Naming

**Severity: HIGH**

Two sets of constants with conflicting indexing schemes:

| Constant | File | Index Type | Usage |
|----------|------|------------|-------|
| `GRIEVANCE_COLS` | 01_Core.gs:894 | 1-indexed | `getRange()` calls |
| `GRIEVANCE_COLUMNS` | 01_Core.gs:977 | 0-indexed | Array access |
| `MEMBER_COLS` | 01_Core.gs:829 | 1-indexed | `getRange()` calls |
| `MEMBER_COLUMNS` | 01_Core.gs:1073 | 0-indexed | Array access |

**Recommendation:** Consolidate to a single naming scheme with clear documentation:
```javascript
// Single source of truth
var COLUMNS = {
  GRIEVANCE: {
    GRIEVANCE_ID: { sheet: 1, array: 0 },  // Explicit about both indices
    FIRST_NAME: { sheet: 2, array: 1 },
    // ...
  }
};
```

### 2.2 Magic Numbers

**Severity: MEDIUM**

46+ occurrences of hardcoded time calculations:

```javascript
// Current (scattered throughout code)
new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000)  // 30 days
new Date(incidentDate.getTime() + 21 * 24 * 60 * 60 * 1000)  // 21 days
```

**Recommendation:** Create a constants file:
```javascript
var TIME_CONSTANTS = {
  MS_PER_DAY: 24 * 60 * 60 * 1000,
  DAYS: {
    FILING_DEADLINE: 21,
    STEP1_RESPONSE: 30,
    STEP2_RESPONSE: 10,
    APPEAL_WINDOW: 90
  }
};

// Usage
new Date(now.getTime() - TIME_CONSTANTS.DAYS.STEP1_RESPONSE * TIME_CONSTANTS.MS_PER_DAY)
```

### 2.3 Functions with Excessive Parameters

**Severity: MEDIUM**

| File | Function | Parameters |
|------|----------|------------|
| `07_DevTools.gs:972` | `generateSingleMemberRow` | 20 parameters |
| `07_DevTools.gs:1171` | `generateSingleGrievanceRow` | 13 parameters |

**Recommendation:** Use object parameters:
```javascript
// Instead of 20 parameters
function generateSingleMemberRow(config) {
  var { memberId, firstName, lastName, jobTitle, ...rest } = config;
  // ...
}
```

### 2.4 Duplicate Function Names

**Severity: HIGH**

34 function names are duplicated across files, including:
- `handleError`
- `handleSearch`
- `displayResults`
- `render`
- `loadData`
- `importMembers`
- `clearAll`
- `saveSettings`
- `showError`
- `showTab`

**Recommendation:** Establish naming conventions with module prefixes or namespaces.

### 2.5 Low Error Handling Coverage

**Statistics:**
- Total functions: 1,067
- Try-catch blocks: 171 (16% coverage)

**Recommendation:** Wrap all sheet operations in try-catch blocks with consistent error handling.

---

## 3. Architecture Recommendations

### 3.1 Implement Data Access Layer

**Current Problem:** 406 direct calls to `SpreadsheetApp.getActiveSpreadsheet()` scattered throughout the codebase.

**Recommendation:** Create a centralized data access layer:

```javascript
// New file: 00_DataAccess.gs
var DataAccess = (function() {
  var _ss = null;

  function getSpreadsheet() {
    if (!_ss) {
      _ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    return _ss;
  }

  function getSheet(sheetName) {
    return getSpreadsheet().getSheetByName(sheetName);
  }

  function batchGetValues(sheetName, ranges) {
    // Optimized batch retrieval
  }

  return {
    getSpreadsheet: getSpreadsheet,
    getSheet: getSheet,
    batchGetValues: batchGetValues
  };
})();
```

### 3.2 Consolidate Trigger Handlers

**Current Problem:** 6 different `onEdit` variants defined across modules (only one can be active):

| File | Function |
|------|----------|
| `10_Main.gs:73` | `onEdit` (main) |
| `06_Maintenance.gs:3294` | `onEditWithAuditLogging` |
| `07_DevTools.gs:2044` | `onEditValidation` |
| `08_SheetUtils.gs:429` | `onEditMultiSelect` |
| `08_SheetUtils.gs:3059` | `onEditAudit` |
| `09_Dashboards.gs:2183` | `onEditAutoSync` |

**Recommendation:** Single dispatcher pattern:
```javascript
// 10_Main.gs
function onEdit(e) {
  TriggerRouter.dispatch('edit', e);
}

var TriggerRouter = {
  handlers: {
    edit: [
      { name: 'validation', fn: handleEditValidation },
      { name: 'multiSelect', fn: handleEditMultiSelect },
      { name: 'audit', fn: handleEditAudit },
      { name: 'autoSync', fn: handleEditAutoSync }
    ]
  },
  dispatch: function(event, e) {
    this.handlers[event].forEach(function(handler) {
      try {
        handler.fn(e);
      } catch (err) {
        Logger.log('Handler ' + handler.name + ' failed: ' + err);
      }
    });
  }
};
```

### 3.3 Break Down Large Files

**Current State:**

| File | Lines | Functions | Recommendation |
|------|-------|-----------|----------------|
| `04_UIService.gs` | 7,127 | 197 | Split into 4 files |
| `10_Code.gs` | 5,406 | 50 | Split into 3 files |
| `08_SheetUtils.gs` | 4,442 | 90 | Split into 4 files |

**Recommended Structure for UIService:**
- `04_UIService.gs` - Dialog orchestration (keep ~2000 lines)
- `04a_UIComponents.gs` - HTML generation helpers
- `04b_UIThemes.gs` - Theme and styling logic
- `04c_UISearch.gs` - Search interface logic

### 3.4 Reduce Module Coupling

**Current Coupling Analysis:**

```
10_Main.gs ──────┬──> 02_DataManagers.gs ──> 06_Maintenance.gs
                 ├──> 05_Integrations.gs ──> 02_DataManagers.gs
                 └──> 06_Maintenance.gs

All modules ────────> 06_Maintenance.gs (audit logging)
```

**Recommendation:** Implement event-driven architecture for cross-cutting concerns:
```javascript
var EventBus = {
  listeners: {},
  on: function(event, callback) {
    if (!this.listeners[event]) this.listeners[event] = [];
    this.listeners[event].push(callback);
  },
  emit: function(event, data) {
    (this.listeners[event] || []).forEach(function(cb) { cb(data); });
  }
};

// Usage in DataManagers
function addMember(memberData) {
  // ... add member logic
  EventBus.emit('member:created', { memberId: newId });
}

// Maintenance module subscribes
EventBus.on('member:created', function(data) {
  logAuditEvent('MEMBER_CREATED', data.memberId);
});
```

---

## 4. Performance Optimizations

### 4.1 Batch Sheet Operations

**Current Problem (05_Integrations.gs:796-806):**
```javascript
// 11 separate API calls for one row
var data = {
  grievanceId: sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue(),
  name: sheet.getRange(row, GRIEVANCE_COLS.FIRST_NAME).getValue() + ' ' +
        sheet.getRange(row, GRIEVANCE_COLS.LAST_NAME).getValue(),
  // ... 8 more getValue() calls
};
```

**Recommendation:**
```javascript
// Single API call
var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
var data = {
  grievanceId: rowData[GRIEVANCE_COLUMNS.GRIEVANCE_ID],
  name: rowData[GRIEVANCE_COLUMNS.FIRST_NAME] + ' ' + rowData[GRIEVANCE_COLUMNS.LAST_NAME],
  // ... use array indices
};
```

### 4.2 Cache Spreadsheet Reference

**Current:** New spreadsheet reference obtained on every operation (406 times)

**Recommendation:** Use module-level caching with invalidation:
```javascript
var SheetCache = {
  _ss: null,
  _sheets: {},

  getSpreadsheet: function() {
    if (!this._ss) {
      this._ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    return this._ss;
  },

  getSheet: function(name) {
    if (!this._sheets[name]) {
      this._sheets[name] = this.getSpreadsheet().getSheetByName(name);
    }
    return this._sheets[name];
  },

  invalidate: function() {
    this._ss = null;
    this._sheets = {};
  }
};
```

---

## 5. Testing Improvements

### 5.1 Current State

| Metric | Value |
|--------|-------|
| Total functions | ~796 |
| Unit tests | ~27 |
| Coverage | ~3.4% |
| Test framework | Custom (non-standard) |

### 5.2 Recommendations

**Priority 1: Test Critical Business Logic**
- `addMember()`, `updateMember()` - Member CRUD
- `createGrievance()`, `advanceGrievanceStep()` - Grievance workflow
- `calculateDeadlines()` - Deadline calculations
- `doGet()` - Web app entry point

**Priority 2: Implement Mock Framework**
```javascript
// Mock for SpreadsheetApp
var MockSpreadsheetApp = {
  _sheets: {},
  getActiveSpreadsheet: function() {
    return {
      getSheetByName: function(name) {
        return MockSpreadsheetApp._sheets[name];
      }
    };
  },
  setMockSheet: function(name, data) {
    this._sheets[name] = new MockSheet(data);
  }
};
```

**Priority 3: Target 70% Coverage**
- Focus on files with business logic: `02_DataManagers.gs`, `10_Code.gs`
- Add integration tests for trigger handlers
- Add validation tests for all input functions

---

## 6. Documentation Updates

### 6.1 Outdated References

**DEVELOPER_GUIDE.md** references modules that no longer exist:
- `04a_MenuBuilder.gs` (consolidated into `03_UIComponents.gs`)
- `04b_ThemeService.gs` (consolidated into `03_UIComponents.gs`)
- `08a_SheetCreation.gs`, `08b_DataValidation.gs`, `08c_SearchEngine.gs`, `08d_ChartBuilder.gs` (all consolidated)

**Recommendation:** Update module table to reflect actual 13-file architecture.

### 6.2 Missing JSDoc Coverage

Files needing JSDoc improvements:
- `01_Core.gs` - Utility functions lack `@param` and `@returns`
- `06_Maintenance.gs` - Cache functions need documentation
- `10_Main.gs` - Trigger handlers need event parameter documentation

### 6.3 Missing API Documentation

**Recommendation:** Create `API_REFERENCE.md` with:
- Function signatures for all public functions
- Parameter types and descriptions
- Return value documentation
- Usage examples

---

## 7. Build System Issues

### 7.1 ESLint Configuration Incompatibility

**Problem:** The project uses ESLint 8.57.0 with a `.eslintrc.js` file, but running `npm run lint` fails because ESLint 9.x (installed) requires `eslint.config.js` format.

**Evidence:**
```
ESLint: 9.39.2
ESLint couldn't find an eslint.config.(js|mjs|cjs) file.
```

**Recommendation:** Either:
1. Pin ESLint to version 8.x in `package.json`:
   ```json
   "eslint": "^8.57.0"
   ```
2. Or migrate to flat config format (`eslint.config.js`)

### 7.2 ESLint Rules Too Lenient

Almost all rules are turned off in `.eslintrc.js`. Consider gradually enabling:
- `no-undef` - Catch undefined variables
- `no-unused-vars` - Remove dead code
- `no-redeclare` - Prevent accidental redeclarations
- `eqeqeq` - Enforce strict equality

---

## 8. Prioritized Action Items

### Immediate (Security - Week 1)

| Priority | Issue | File(s) | Effort |
|----------|-------|---------|--------|
| P0 | Fix XSS vulnerabilities | Multiple | 1-2 days |
| P0 | Add input validation to doGet | 05_Integrations.gs | 1 day |
| P0 | Implement access control | 05_Integrations.gs | 2-3 days |
| P1 | Fix formula injection | 08_SheetUtils.gs, 12_Features.gs | 1 day |
| P1 | Remove PII from logs | Multiple | 0.5 days |

### Short-term (Code Quality - Week 2-3)

| Priority | Issue | File(s) | Effort |
|----------|-------|---------|--------|
| P1 | Consolidate constant naming | 01_Core.gs | 2 days |
| P1 | Fix ESLint configuration | .eslintrc.js / package.json | 0.5 days |
| P2 | Extract magic numbers | Multiple | 1 day |
| P2 | Resolve duplicate functions | Multiple | 2-3 days |
| P2 | Add error handling | Multiple | 2 days |

### Medium-term (Architecture - Month 1-2)

| Priority | Issue | Effort |
|----------|-------|--------|
| P2 | Implement Data Access Layer | 3-5 days |
| P2 | Consolidate trigger handlers | 2 days |
| P3 | Break down large files | 5-7 days |
| P3 | Implement event bus | 3 days |

### Long-term (Quality - Month 2-3)

| Priority | Issue | Effort |
|----------|-------|--------|
| P3 | Increase test coverage to 70% | 2-3 weeks |
| P3 | Update documentation | 1 week |
| P3 | Create API reference | 1 week |

---

## Summary

This codebase has solid functionality and recent improvements in modularization, but requires immediate attention to critical security vulnerabilities (XSS, missing access control, formula injection). Code quality improvements should focus on consolidating constants, reducing duplication, and improving error handling. Long-term architectural improvements will significantly enhance maintainability and testability.

The project benefits from excellent technical debt management (only 1 TODO) and reasonable documentation coverage. With the recommended changes, this can become a highly maintainable and secure application.

---

*Generated by Claude Code Review*
