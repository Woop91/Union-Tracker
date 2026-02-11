# Code Review: 509 Strategic Command Center

**Date:** 2026-02-11
**Scope:** Full codebase review (~49,889 lines across 16 source files)
**Lint Status:** ESLint passes clean
**Test Status:** 871 tests passing across 17 suites

---

## Critical Issues

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
| Critical | 2 | XSS vulnerabilities, string split bug |
| High | 3 | Dual constant systems, hardcoded indices, shared undo state |
| Medium | 4 | Hardcoded sheet names, dev tools in prod, weak anonymization hash, column adjacency assumption |
| Low | 4 | ID sequence pollution, character stripping, trend ordering, missing CSP |
