# Code Review: 509 Strategic Command Center

**Date:** 2026-02-11
**Scope:** Full codebase review (~49,889 lines across 16 source files)
**Lint Status:** ESLint passes clean
**Test Status:** 871 tests passing across 17 suites

---

## Critical Issues

### 0. `setupHiddenSheets()` References 4 Undefined Properties — Main Setup Broken

**File:** `src/08_SheetUtils.gs:207-215`

The `setupHiddenSheets(ss)` function is called by `CREATE_509_DASHBOARD()` (the main first-time setup) and iterates over 6 hidden sheet configurations. However, **4 of 6 property names don't exist** in the `HIDDEN_SHEETS` constant:

```javascript
// setupHiddenSheets() references:          HIDDEN_SHEETS constant defines:
HIDDEN_SHEETS.CALC_MEMBERS    → undefined   // (not defined)
HIDDEN_SHEETS.CALC_GRIEVANCES → undefined   // (not defined)
HIDDEN_SHEETS.CALC_DEADLINES  → undefined   // (not defined)
HIDDEN_SHEETS.CALC_STATS      → '_Dashboard_Calc'      ✓
HIDDEN_SHEETS.CALC_SYNC       → undefined   // (not defined)
HIDDEN_SHEETS.CALC_FORMULAS   → '_Grievance_Formulas'   ✓
```

When `name` is `undefined`, `ss.getSheetByName(undefined)` returns `null`, then `ss.insertSheet(undefined)` is called. This will either throw an error (crashing `CREATE_509_DASHBOARD()`) or create sheets with auto-generated names that no other code references.

**Impact:** The main dashboard setup function (`CREATE_509_DASHBOARD()`) will fail or produce broken hidden sheets. 4 hidden calculation sheets (Members stats, Grievances stats, Deadlines, Sync) won't be properly created. The setup functions for these sheets exist and are correct (`setupCalcMembersSheet`, `setupCalcGrievancesSheet`, `setupCalcDeadlinesSheet`, `setupCalcSyncSheet` at lines 4092-4435) — they just can't be reached because the sheet names are undefined.

**Note:** A parallel system exists — `setupAllHiddenSheets()` (line 3925) creates 7 *different* hidden sheets (`_Grievance_Calc`, `_Member_Lookup`, `_Steward_Contact_Calc`, `_Steward_Performance_Calc`, `_Checklist_Calc`) using properly named functions. But this function is NOT called by `CREATE_509_DASHBOARD()`.

**Fix:** Add the missing properties to `HIDDEN_SHEETS`:
```javascript
var HIDDEN_SHEETS = {
  CALC_MEMBERS: '_Calc_Members',         // ADD
  CALC_GRIEVANCES: '_Calc_Grievances',   // ADD
  CALC_DEADLINES: '_Calc_Deadlines',     // ADD
  CALC_SYNC: '_Calc_Sync',              // ADD
  CALC_STATS: '_Dashboard_Calc',
  CALC_FORMULAS: '_Grievance_Formulas',
  // ... existing entries ...
};
```

Or call `setupAllHiddenSheets()` from `CREATE_509_DASHBOARD()` in addition to (or instead of) `setupHiddenSheets()`.

---

### 1. XSS Vulnerabilities — User Data Injected into HTML Without Escaping

Multiple HTML-generating functions interpolate user-controlled data directly into HTML strings without using `escapeHtml()`. While `escapeHtml()` exists in `00_Security.gs` and is used in some places (e.g., email templates, survey dashboard), it is missing in several critical locations.

**Affected locations:**

| File | Function | Line | Unescaped Data |
|------|----------|------|----------------|
| `10_Main.gs` | `getEditGrievanceFormHtml()` | ~1148 | Grievance description inserted into HTML form value |
| `10_Main.gs` | `getNewGrievanceFormHtml()` | ~1400 | Member names in `<option>` elements |
| `03_UIComponents.gs` | Quick actions HTML | various | `memberId` in `onclick` handlers |
| `03_UIComponents.gs` | `showMobileGrievanceList()` | various | `g.memberName` rendered in search results |
| `05_Integrations.gs` | `openGrievanceFolder()` | 364-375 | `folderUrl` and `result.folderUrl` injected into `<script>` tags |
| `11_CommandHub.gs` | `getErrorPageHtml_()` | 3544 | `message` param inserted into HTML |
| `11_CommandHub.gs` | `getMemberPortalHtml_()` | 3395 | `profile.firstName` in welcome text |
| `11_CommandHub.gs` | `getPublicPortalHtml_()` | 3414, 3508 | Steward names in portal HTML |
| `12_Features.gs` | `buildReminderDialogHtml_()` | 2308-2309 | `grievanceId`, `reminders.memberName`, `reminders.status` |

**Risk:** An attacker who can control a grievance description, member name, or folder URL could inject JavaScript that executes in the context of the Google Apps Script HtmlService dialog. In the `openGrievanceFolder()` case, a malicious URL could break out of the string and execute arbitrary script.

**Recommendation:** Apply `escapeHtml()` to all user-controlled values before inserting them into HTML. For values inserted into `<script>` blocks, use `JSON.stringify()` instead.

---

### 2. String Split Bug in `importMembersFromText()`

**File:** `src/10_Main.gs:1572`

```javascript
const lines = text.split('\\n').filter(line => line.trim());
```

The `'\\n'` is a two-character string literal (backslash + n), not a newline character. This means `text.split('\\n')` looks for the literal characters `\n` in the text, not actual line breaks. If `text` comes from a textarea, it will contain real newline characters (`\n`), and the split will produce a single-element array containing the entire text.

**Fix:** Change to `text.split('\n')`.

---

## High Priority Issues

### 3. Dual Column Constant Systems Create Maintenance Risk

The codebase maintains two parallel column indexing systems:

- **0-indexed** (`GRIEVANCE_COLUMNS`, `MEMBER_COLUMNS`) — for array access via `data[row][col]`
- **1-indexed** (`GRIEVANCE_COLS`, `MEMBER_COLS`) — for sheet operations via `sheet.getRange(row, col)`

While functions using `GRIEVANCE_COLUMNS` do correctly convert with `+ 1` for `getRange()` calls, having two systems introduces ongoing risk. New contributors must know which to use and when to add `+ 1` or `- 1`.

**Examples of the dual usage:**
- `updateGrievanceFolderLink()` in `05_Integrations.gs:341-342` uses `GRIEVANCE_COLUMNS` (0-indexed) with `+ 1` for getRange
- `setupFolderForSelectedGrievance()` in the same file uses `GRIEVANCE_COLS` (1-indexed) with `- 1` for array access
- `advancedSearch()` in `08_SheetUtils.gs:865` uses `MEMBER_COLUMNS` while `getDesktopSearchData()` at line 647 uses `MEMBER_COLS`

**Recommendation:** Consolidate to a single constant system. The 1-indexed system (`GRIEVANCE_COLS`/`MEMBER_COLS`) with `- 1` for array access is the more common pattern. Deprecate and remove the 0-indexed versions.

---

### 4. Hardcoded Column Indices in Satisfaction Stats

**File:** `src/11_CommandHub.gs:3024-3025`

```javascript
var trustVal = parseFloat(data[i][7]); // SATISFACTION_COLS.Q7_TRUST_UNION - 1
var satVal = parseFloat(data[i][6]);   // SATISFACTION_COLS.Q6_SATISFIED_REP - 1
```

These use hardcoded magic numbers instead of the `SATISFACTION_COLS` constants. If column order changes in the satisfaction survey sheet, these will silently read the wrong data.

**Recommendation:** Use `SATISFACTION_COLS` constants with `- 1` for array indexing, matching the pattern used elsewhere.

---

### 5. Shared State in `ScriptProperties` — Multi-User Conflicts

**File:** `src/06_Maintenance.gs`

The undo/redo system and cache layer store state in `ScriptProperties`, which is shared across all users of the spreadsheet. This means:

- One user's undo history is visible/usable by another user
- A user could undo another user's changes
- Cache entries are shared, which could cause stale data for concurrent users

**Recommendation:** For undo history, consider using `PropertiesService.getUserProperties()` instead of `ScriptProperties` to isolate per-user state. For caching, `CacheService.getScriptCache()` is already shared by design which is acceptable for read caching, but undo/redo should be user-scoped.

---

## Medium Priority Issues

### 6. Hardcoded Sheet Names in Config Objects

Several configuration objects hardcode sheet names instead of referencing the `SHEETS` constant:

| File | Config | Hardcoded Values |
|------|--------|------------------|
| `11_CommandHub.gs` | `COMMAND_CENTER_CONFIG` | `'Grievance Log'`, `'Member Directory'` |
| `12_Features.gs` | `EXTENSION_CONFIG` | `'Member Directory'`, `'Grievance Log'` |
| `12_Features.gs` | `LOOKER_CONFIG.ALLOWED_SOURCES` | `['Member Directory', 'Grievance Log', 'Member Satisfaction']` |

These will break if sheet names are ever changed in the `SHEETS` constant without updating these config objects.

**Recommendation:** Reference `SHEETS` constant values where possible, or add comments documenting the dependency.

---

### 7. `07_DevTools.gs` — "DELETE THIS FILE BEFORE PRODUCTION"

The file header contains the comment "DELETE THIS FILE BEFORE PRODUCTION". It contains functions like `NUKE_DATABASE()`, `SEED_MEMBERS()`, `SEED_GRIEVANCES()` that create/destroy test data. These are exposed via the Apps Script menu and could be accidentally triggered by end users.

**Recommendation:** If this code is in production, either remove the file or add access control checks (e.g., verify the user is a developer/admin) before allowing destructive operations. The `NUKE_DATABASE()` function deletes all data from all sheets including the audit log.

---

### 8. Weak Hash Function for Anonymized Data

**File:** `src/12_Features.gs:3541-3552`

```javascript
function generateAnonHash_(id) {
  const salt = 'anon509data';
  const combined = salt + String(id);
  let hash = 0;
  for (let i = 0; i < combined.length; i++) {
    const char = combined.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return 'A' + Math.abs(hash).toString(36).toUpperCase().substring(0, 8);
}
```

This is a djb2-like hash with a static hardcoded salt. Problems:
- The hash space is limited to 32-bit integers, making collisions likely with many records
- The static salt means anyone with the source code can brute-force member IDs (which are typically sequential/predictable) to reverse the anonymization
- The `substring(0, 8)` further reduces the hash space

**Recommendation:** Use `Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + id)` for proper anonymization. Use a configurable salt stored in `ScriptProperties` rather than hardcoded in source.

---

### 9. `setGrievanceReminder()` Assumes Contiguous Columns

**File:** `src/12_Features.gs:2046-2047`

```javascript
const updates = [[parsedDate || '', reminderNote || '']];
sheet.getRange(rowIndex, dateCol, 1, 2).setValues(updates);
```

This writes a 2-column range starting from `dateCol`, assuming the note column is immediately adjacent to the date column (`dateCol + 1`). If the sheet schema changes such that `REMINDER_1_NOTE` is not at `REMINDER_1_DATE + 1`, this will write the note to the wrong column.

**Recommendation:** Write each value separately using `noteCol`, or verify the assumption with an assertion.

---

## Low Priority Issues

### 10. `verifyIDGenerationEngine()` Increments Production Sequence

**File:** `src/11_CommandHub.gs:336`

The `verifyIDGenerationEngine()` function creates test ID sequence entries, which increment the production sequence counter but the test entries are never cleaned up. Running this verification repeatedly will cause ID gaps.

---

### 11. Checklist `caseId` Sanitization Strips Characters

**File:** `src/12_Features.gs:957` (approximately)

The checklist dialog sanitizes `caseId` with `replace(/['"\\<>&]/g, '')` which strips characters. While this prevents injection, if a grievance ID legitimately contained these characters (unlikely but possible with `&`), it would silently fail to match.

---

### 12. `getSecureSatisfactionStats_()` Trend Analysis Window

**File:** `src/11_CommandHub.gs:3044-3047`

The trend analysis compares the last 10 trust scores vs the previous 10. This treats the data as time-ordered based on array position, but there's no guarantee the rows are sorted chronologically. If rows are reordered or inserted, the trend calculation becomes meaningless.

---

### 13. Portal HTML Lacks Content Security Policy

The member portal and public portal HTML (`getMemberPortalHtml_()`, `getPublicPortalHtml_()`) load external resources from `fonts.googleapis.com` but don't set a Content-Security-Policy meta tag. While GAS HtmlService provides some sandboxing, adding CSP headers would add defense in depth.

---

## Sheet/Tab Population Analysis

### Sheets Created and Populated by `CREATE_509_DASHBOARD()`

| # | Sheet Name | Created By | Populated? |
|---|-----------|-----------|------------|
| 1 | `Config` | `createConfigSheet()` | Yes — 52 columns of dropdown sources, defaults, section headers |
| 2 | `Member Directory` | `createMemberDirectory()` | Yes — headers, validation, conditional formatting, checkboxes, column groups, filters |
| 3 | `Grievance Log` | `createGrievanceLog()` | Yes — headers, validation, formatting, checkboxes, column groups, deadline heatmap |
| 4 | `Case Checklist` | `getOrCreateChecklistSheet()` | Yes — 12-column headers, dropdowns, checkboxes |
| 5 | `📊 Member Satisfaction` | `createSatisfactionSheet()` | Yes — 67 survey question headers + 11 section average formula columns |
| 6 | `💡 Feedback & Development` | `createFeedbackSheet()` | Yes — headers, dropdowns, conditional formatting, metrics section |
| 7 | `Function Checklist` | `createFunctionChecklistSheet_()` | Yes — 100+ menu items with function names and descriptions |
| 8 | `📚 Getting Started` | `createGettingStartedSheet()` | Yes — styled setup guide with instructions |
| 9 | `❓ FAQ` | `createFAQSheet()` | Yes — 15+ Q&A items with styling |
| 10 | `📖 Config Guide` | `createConfigGuideSheet()` | Yes — instructional guide with column reference |

### Hidden Sheets Created by `setupHiddenSheets()` (called from `CREATE_509_DASHBOARD`)

| Sheet Name | Status | Notes |
|-----------|--------|-------|
| `_Dashboard_Calc` (via `CALC_STATS`) | **Works** | Dashboard-wide stats with formulas |
| `_Grievance_Formulas` (via `CALC_FORMULAS`) | **Works** | 24 columns of VLOOKUP/deadline formulas |
| `CALC_MEMBERS` sheet | **BROKEN** — property undefined | Would contain member statistics and lookups |
| `CALC_GRIEVANCES` sheet | **BROKEN** — property undefined | Would contain grievance aggregations |
| `CALC_DEADLINES` sheet | **BROKEN** — property undefined | Would contain deadline calculations and alerts |
| `CALC_SYNC` sheet | **BROKEN** — property undefined | Would contain cross-sheet sync and consistency checks |

### Hidden Sheets Created by `setupAllHiddenSheets()` (separate function, NOT called during setup)

| Sheet Name | Populated? | Notes |
|-----------|------------|-------|
| `_Grievance_Calc` | Yes | 7 columns — member-level grievance stats with ARRAYFORMULA |
| `_Grievance_Formulas` | Yes | 24 columns — VLOOKUP/deadline formulas (overlaps with System A) |
| `_Member_Lookup` | Yes | 7 columns — member data lookups via VLOOKUP |
| `_Steward_Contact_Calc` | Yes | 5 columns — steward contact metrics |
| `_Dashboard_Calc` | Yes | 15+ metrics with COUNTIF formulas (overlaps with System A) |
| `_Steward_Performance_Calc` | Yes | 10 columns — weighted steward performance scores |
| `_Checklist_Calc` | Yes | 8 columns — checklist progress tracking |

### Sheets Defined in `SHEETS` Constant but NOT Created During Setup

| Sheet Name | Why | Risk |
|-----------|-----|------|
| `📋 Features Reference` | Creation function exists (`createFeaturesReferenceSheet`) but not called from `CREATE_509_DASHBOARD()` — only available via menu | Low — menu item works if user clicks it, but won't exist after initial setup |
| `💼 Dashboard` | Deprecated — listed in `reorderSheetsToStandard()` but never created | None — gracefully skipped during reorder |
| `Test Results` | Only created by test runner, not during setup | None — `viewTestResults()` shows alert if missing |
| `📅 Meeting Attendance` | Labeled "Optional source sheets" | None — not referenced by core functionality |
| `🤝 Volunteer Hours` | Labeled "Optional source sheets" | None — not referenced by core functionality |

### Sheets Missing from `reorderSheetsToStandard()`

These sheets are created during setup but not included in the sheet ordering logic:
- `Case Checklist` — will end up in default position
- `📋 Features Reference` — if manually created via menu, won't be ordered

### Looker Integration Sheets (Created by Separate Setup Functions)

| Sheet Name | Created By | Populated? | Hidden? |
|-----------|-----------|------------|---------|
| `_Looker_Grievances` | `setupLookerIntegration()` | Yes — 31 headers + data refresh | Yes |
| `_Looker_Members` | `setupLookerIntegration()` | Yes — 22 headers + data refresh | Yes |
| `_Looker_Satisfaction` | `setupLookerIntegration()` | Yes — 30 headers + data refresh | Yes |
| `_Looker_Anon_Grievances` | `setupLookerAnonIntegration()` | Yes — 24 headers + anonymized data | Yes |
| `_Looker_Anon_Members` | `setupLookerAnonIntegration()` | Yes — 15 headers + anonymized data | Yes |
| `_Looker_Anon_Satisfaction` | `setupLookerAnonIntegration()` | Yes — 22 headers + anonymized data | Yes |

These are correctly implemented and populate properly when their respective setup functions are called.

### Audit Log Sheet

| Sheet Name | Created By | Populated? | Hidden? |
|-----------|-----------|------------|---------|
| `_Audit_Log` | `setupAuditLogSheet()` (on demand) | Headers only — data added by `logAuditEvent()` trigger | Yes |

---

## Positive Observations

- **Test coverage is solid:** 871 tests across 17 suites all passing
- **ESLint passes clean** with no warnings
- **Security fundamentals are in place:** PIN hashing with SHA-256 + salt, rate limiting, session tokens, audit logging, IDOR protection on web app endpoints
- **PII protection is thoughtful:** The `safetyValveScrub()` function, anonymized Looker sheets, and PII-free exports demonstrate privacy awareness
- **Batch operations are well-optimized:** Functions consistently use `setValues()` for batch writes and minimize API calls
- **Build/deploy pipeline exists:** `build.js` concatenates modules into a distributable file
- **Sabotage detection:** The `onEdit` trigger monitors for mass cell deletion (>15 cells) which is a practical safeguard
- **`MultiSelectDialog.html` is clean:** Uses DOM API (`textContent`) for rendering which is inherently XSS-safe

---

## Summary

| Severity | Count | Key Theme |
|----------|-------|-----------|
| Critical | 3 | Broken hidden sheet setup (4 undefined properties), XSS vulnerabilities, string split bug |
| High | 3 | Dual constant systems, hardcoded indices, shared undo state |
| Medium | 4 | Hardcoded sheet names, dev tools in prod, weak anonymization hash, column adjacency assumption |
| Low | 4 | ID sequence pollution, character stripping, trend ordering, missing CSP |
