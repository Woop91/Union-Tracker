# 509 Dashboard Developer Guide

This guide provides an overview of the codebase architecture, development workflow, and best practices for contributing to the Union Steward Dashboard.

## Table of Contents

1. [Project Overview](#project-overview)
2. [Architecture](#architecture)
3. [Module Structure](#module-structure)
4. [Build System](#build-system)
5. [Testing](#testing)
6. [Code Style Guidelines](#code-style-guidelines)
7. [Common Patterns](#common-patterns)
8. [Debugging](#debugging)

---

## Project Overview

The 509 Dashboard is a Google Apps Script (GAS) application for managing union steward activities, grievances, and member information within Google Sheets.

**Key Technologies:**
- Google Apps Script (ES5-compatible JavaScript)
- Google Sheets API
- Node.js build system
- ESLint for code quality

**Version:** 4.4.1

---

## Architecture

### High-Level Structure

```
MULTIPLE-SCRIPS-REPO/
├── src/                    # Source files (.gs)
├── dist/                   # Build output (auto-generated)
├── .github/workflows/      # CI/CD configuration
├── build.js               # Build script
├── package.json           # Dependencies and scripts
├── .eslintrc.js           # Linting configuration
└── DEVELOPER_GUIDE.md     # This file
```

### Modular Design

The codebase follows a modular architecture where each file handles a specific domain:

```
┌─────────────────────────────────────────────────────────────┐
│                      09_Main.gs                              │
│                   (Entry Point / onOpen)                     │
└─────────────────────────────────────────────────────────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        ▼                     ▼                     ▼
┌───────────────┐   ┌───────────────┐   ┌───────────────┐
│  UI Layer     │   │  Data Layer   │   │  Services     │
│  04a/04b/04   │   │  02/03/08     │   │  05/06/14     │
└───────────────┘   └───────────────┘   └───────────────┘
        │                     │                     │
        └─────────────────────┼─────────────────────┘
                              ▼
                    ┌───────────────┐
                    │ 01_Constants  │
                    │   (Globals)   │
                    └───────────────┘
```

---

## Module Structure

### Core Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| `01_Constants.gs` | Global constants and configuration | `SHEETS`, `COLORS`, column indices |
| `02_MemberManager.gs` | Member CRUD operations | `getMemberById`, `updateMember`, `addMember` |
| `03_GrievanceManager.gs` | Grievance handling | `createGrievance`, `updateGrievance`, `getGrievanceStats` |

### UI Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| `04a_MenuBuilder.gs` | Menu creation | `createDashboardMenu`, `navigateToSheet` |
| `04b_ThemeService.gs` | Theme management | `APPLY_SYSTEM_THEME`, `toggleDarkMode`, `applyZebraStripes` |
| `04_UIService.gs` | Dialogs and panels | `showMemberDialog`, `showGrievanceForm` |

### Maintenance Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| `06a_Diagnostics.gs` | System diagnostics | `DIAGNOSE_SETUP`, `REPAIR_DASHBOARD` |
| `06b_CacheManager.gs` | Caching layer | `getCachedData`, `invalidateCache`, `warmUpCaches` |
| `06c_UndoManager.gs` | Undo/redo system | `recordAction`, `undoLastAction`, `redoLastAction` |
| `06_Maintenance.gs` | General maintenance | `cleanupOldData`, `optimizeSheets` |

### Data Modules

| File | Purpose | Key Functions |
|------|---------|---------------|
| `08a_SheetCreation.gs` | Sheet setup | `CREATE_509_DASHBOARD`, `createGrievanceLogSheet` |
| `08b_DataValidation.gs` | Validation rules | `setupDataValidations`, `setDropdownValidation` |
| `08c_SearchEngine.gs` | Search functionality | `searchDashboard`, `advancedSearch`, `navigateToSearchResult` |
| `08_Code.gs` | Core utilities | Various helper functions |

### Feature Modules

| File | Purpose |
|------|---------|
| `10_CommandCenter.gs` | Admin command center |
| `11_SecureMemberDashboard.gs` | Member-facing dashboard |
| `12_ChecklistManager.gs` | Task checklists |
| `13_DynamicEngine.gs` | Dynamic content generation |
| `14_LookerIntegration.gs` | Looker Studio integration |
| `15_TestFramework.gs` | Unit testing framework |

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

# Lint + Build (test command)
npm run test

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
4. **Output**: Single `dist/ConsolidatedDashboard.gs` file

### Build Order

Files must be concatenated in dependency order:

```javascript
const BUILD_ORDER = [
  '01_Constants.gs',        // Must be first - defines globals
  '02_MemberManager.gs',
  '03_GrievanceManager.gs',
  '04a_MenuBuilder.gs',
  '04b_ThemeService.gs',
  '04_UIService.gs',
  '05_Integrations.gs',
  '06a_Diagnostics.gs',
  '06b_CacheManager.gs',
  '06c_UndoManager.gs',
  '06_Maintenance.gs',
  '07_DevTools.gs',
  '08a_SheetCreation.gs',
  '08b_DataValidation.gs',
  '08c_SearchEngine.gs',
  '08_Code.gs',
  '09_Main.gs',
  '10_CommandCenter.gs',
  '11_SecureMemberDashboard.gs',
  '12_ChecklistManager.gs',
  '13_DynamicEngine.gs',
  '14_LookerIntegration.gs',
  '15_TestFramework.gs'     // Must be last - tests other modules
];
```

---

## Testing

### Test Framework

The project includes a custom test framework in `15_TestFramework.gs`:

```javascript
// Run all tests
runAllTests();

// Show test dashboard UI
showTestDashboard();
```

### Writing Tests

```javascript
function test_exampleFunction() {
  var result = exampleFunction(42);

  assert.equals(result, 42, 'Should return input value');
  assert.isTrue(result > 0, 'Should be positive');
  assert.isDefined(result, 'Should not be undefined');
}
```

### Assert Methods

| Method | Description |
|--------|-------------|
| `assert.isTrue(val, msg)` | Value is truthy |
| `assert.isFalse(val, msg)` | Value is falsy |
| `assert.equals(a, b, msg)` | Strict equality |
| `assert.isDefined(val, msg)` | Not undefined |
| `assert.isArray(val, msg)` | Is an array |
| `assert.contains(arr, val, msg)` | Array contains value |
| `assert.throws(fn, msg)` | Function throws error |

---

## Code Style Guidelines

### General Rules

1. **Use `var` instead of `let`/`const`** - GAS uses ES5
2. **No arrow functions** - Use `function` keyword
3. **JSDoc comments** - Document all public functions
4. **Semicolons required** - End statements with `;`

### Naming Conventions

```javascript
// Functions: camelCase
function getUserById(userId) { }

// Menu-callable functions: UPPER_SNAKE_CASE
function CREATE_509_DASHBOARD() { }

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
| "Sheet not found" | Check `SHEETS` constant in `01_Constants.gs` |
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
