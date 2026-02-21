# Column Position Issues Log

Tracking all column-related bugs found and resolved across the codebase.
The core problem: column positions were hardcoded in many places, so
any manual reordering of sheet columns would silently break formulas,
data access, dropdown validation, and bidirectional sync.

---

## Root Cause

When the project was first built, column positions were assumed to be
fixed. Functions referenced columns by literal numbers (`getRange(2, 5)`),
array indices (`row[3]`), VLOOKUP column offsets (`'3,FALSE)'`), and
column letters (`'$E2'`). Over time, header maps and `*_COLS` constants
were introduced, but many call sites were never migrated.

---

## Issue 1: Hardcoded column references in formulas and data access

**Commit:** `a545022` — _fix: make column mapping fully dynamic so
dropdowns always match columns_

**Severity:** HIGH

**Problem:**
Multiple functions across 10+ source files used hardcoded column numbers
instead of the `MEMBER_COLS`, `GRIEVANCE_COLS`, `CONFIG_COLS`, etc.
constants. Categories included:

- VLOOKUP formula strings with literal column indices
- `getRange(row, 5)` instead of `getRange(row, MEMBER_COLS.WORK_LOCATION)`
- `row[3]` instead of `row[MEMBER_COLS.JOB_TITLE - 1]`
- Column letter references like `'$E2'` instead of `getColumnLetter()`

**Files affected (initial sweep):**
- `08d_AuditAndFormulas.gs` — VLOOKUP column indices, formula column letters
- `10a_SheetCreation.gs` — sheet setup formulas
- `08a_SheetSetup.gs` — data validation setup
- `10_Main.gs` — onEdit handlers
- `02_DataManagers.gs` — data row access
- Various dashboard files

**Resolution:**
Replaced all hardcoded references with dynamic constants. VLOOKUP indices
are now computed as offsets from the range start column. Formula column
letters use `getColumnLetter()`. Array access uses `*_COLS.KEY - 1`.

---

## Issue 2: syncDropdownToConfig_() mapping was incomplete

**Commit:** `c033676` — _fix: centralize dropdown map so validation
and sync can never drift apart_

**Severity:** HIGH

**Problem:**
`setupDataValidations()` set up dropdown validation for 10 single-select
Member Directory columns and 4 Grievance Log columns. But
`syncDropdownToConfig_()` (the bidirectional sync that adds custom
dropdown values back to Config) only mapped 5 + 3 of those columns.

Missing from Member Directory sync: `CONTACT_STEWARD`
Missing from Grievance Log sync: `ARTICLES`

Custom values typed into those columns would never be persisted to the
Config sheet, so they would disappear on the next validation refresh.

**Root cause within the root cause:**
The dropdown-to-Config mapping was defined **independently in two
places** — `setupDataValidations()` had one list, `syncDropdownToConfig_()`
had a separate inline map. When new columns were added to validation,
the sync function was not updated.

**Resolution:**
Created `DROPDOWN_MAP` (via `buildDropdownMap_()`) in `01_Core.gs` as a
single source of truth. Both `setupDataValidations()` and
`syncDropdownToConfig_()` now loop over this same definition.
`syncColumnMaps()` rebuilds it when columns shift.

To add a new dropdown column in the future: add one entry to
`buildDropdownMap_()` (single-select) or `buildMultiSelectCols_()`
(multi-select). Validation and sync both update automatically.

---

## Issue 3: Column positions lost between GAS execution contexts

**Commit:** `1e8c0d7` — _fix: persist column positions across execution
contexts via CacheService_

**Severity:** HIGH

**Problem:**
In Google Apps Script, every trigger execution (`onEdit`, `onOpen`, menu
clicks) starts a **brand-new V8 isolate**. Global column constants are
re-initialized from their default (array-order) positions each time.

`syncColumnMaps()` ran inside `onOpen()` and correctly resolved actual
sheet header positions — but those updated values only lived for that
one execution. The next `onEdit()` started fresh with defaults. If a
user had manually inserted or reordered a column, every subsequent
`onEdit()` would silently operate on the wrong columns.

**Resolution:**
- `syncColumnMaps()` now calls `persistColumnMaps_()` to write resolved
  positions to CacheService (6-hour TTL)
- `onEdit()` calls `loadCachedColumnMaps_()` at the top to restore
  positions from cache (~50ms vs ~500ms for 10 sheet header reads)
- `loadCachedColumnMaps_()` rebuilds all derived objects (legacy compat
  maps, `DROPDOWN_MAP`, `MULTI_SELECT_COLS`) when positions differ
- Graceful degradation: if CacheService is unavailable, defaults are
  used (still correct for sheets created by this code)

---

## Issue 4: Scattered hardcoded indices in data access and utilities

**Commit:** `9ad8b64` — _fix: replace remaining hardcoded column
indices with constants_

**Severity:** MEDIUM

**Problem:**
A second audit pass found additional hardcoded column positions that the
first sweep missed:

| File | Line | Issue |
|---|---|---|
| `00_DataAccess.gs` | 360 | `findRow(sheetName, 0, memberId)` — literal `0` instead of `MEMBER_COLUMNS.MEMBER_ID` |
| `00_DataAccess.gs` | 447 | `findRow(sheetName, 0, grievanceId)` — literal `0` instead of `GRIEVANCE_COLUMNS.GRIEVANCE_ID` |
| `05_Integrations.gs` | 1211 | `member[Object.keys(member)[MEMBER_COLUMNS.EMAIL]]` — fragile fallback assuming JS object key order matches column order |
| `06_Maintenance.gs` | 2328 | `getRange(2, 1, ..., 4)` — hardcoded 4-column width; if GRIEVANCE_ID/MEMBER_ID/FIRST_NAME/LAST_NAME moved past column 4, access would go out of bounds |
| `07_DevTools.gs` | 1561 | `getRange(2, 1, ..., 1)` — hardcoded column 1 instead of `MEMBER_COLS.MEMBER_ID` for seeded data cleanup |
| `07_DevTools.gs` | 1568 | `getRange(2, 1, ..., 1)` — same issue for `GRIEVANCE_COLS.GRIEVANCE_ID` |
| `12_Features.gs` | 502 | `data[i][0]` / `data[i][1]` — literal indices instead of `CHECKLIST_COLS.CHECKLIST_ID` / `CHECKLIST_COLS.CASE_ID` |

Also tightened `loadCachedColumnMaps_()` to only overwrite keys that
exist in **both** the cache and the target object, preventing stale keys
from leaking in after code updates.

**Resolution:**
All replaced with the appropriate `*_COLS` constants. Dynamic column
width computed with `Math.max()` where a multi-column range is needed.

---

## Issue 5: DROPDOWN_MAP / MULTI_SELECT_COLS overlap and Open Rate data loss

**Commit:** `0e994f7` — _fix: correct Open Rate dropdown mapping and
column references_

**Severity:** MEDIUM

**Problem:**
Three interrelated issues found during Open Rate column investigation:

1. **DROPDOWN_MAP / MULTI_SELECT_COLS overlap** — `ARTICLES` and
   `ISSUE_CATEGORY` appeared in both `buildDropdownMap_()` (single-select)
   and `buildMultiSelectCols_()` (multi-select). `setupDataValidations()`
   applied single-select validation first, then multi-select overwrote it.
   The duplicate entries risked the wrong validation type being applied
   and caused `syncDropdownToConfig_()` to treat these columns as
   single-select for bidirectional Config sync.

2. **Multi-select columns excluded from Config sync** —
   `syncDropdownToConfig_()` only iterated over `DROPDOWN_MAP`, not
   `MULTI_SELECT_COLS`. After removing the duplicates from
   `DROPDOWN_MAP` (fix #1), custom values typed into multi-select columns
   would stop syncing back to Config. Updated the function to also
   iterate over `MULTI_SELECT_COLS`.

3. **Open Rate 0% data loss** — `mapMemberRow()` used
   `openRate: row[MEMBER_COLS.OPEN_RATE - 1] || ''` which treats `0` as
   falsy in JavaScript, converting a legitimate 0% open rate to an empty
   string. Introduced `numericField_()` helper for null-safe numeric field
   access that preserves `0` values.

**Additional cleanup:**
Corrected inline column letter comments across 5 files — Open Rate is
column T (not S), Volunteer Hours is column U (not T), etc. These
comments were wrong after a prior column insertion but had no functional
impact since the code uses `*_COLS` constants.

**Files affected:**
- `01_Core.gs` — removed ARTICLES/ISSUE_CATEGORY from `buildDropdownMap_()`,
  added `numericField_()`, fixed `mapMemberRow()`
- `10_Main.gs` — updated `syncDropdownToConfig_()` to also sync
  `MULTI_SELECT_COLS`
- `10a_SheetCreation.gs` — corrected column letter comments
- `10d_SyncAndMaintenance.gs` — corrected column letter comments
- `04e_PublicDashboard.gs` — corrected column letter comments
- `05_Integrations.gs` — corrected column letter comments
- `10b_SurveyDocSheets.gs` — corrected column letter comments
- `dist/ConsolidatedDashboard.gs` — consolidated build reflecting all changes

**Resolution:**
Single-select and multi-select dropdown mappings are now mutually
exclusive. Both are synced bidirectionally with Config. Numeric fields
use `numericField_()` to safely preserve zero values.

---

## Issue 6: mapGrievanceRow and export functions lose zero-value numeric fields

**Commit:** _(included in this review cycle)_

**Severity:** MEDIUM

**Problem:**
The same `|| ''` falsy bug fixed for `openRate` in Issue 5 also
affected `daysOpen` and `daysToDeadline` in three locations:

1. **`mapGrievanceRow()`** in `01_Core.gs` (lines 2011, 2013) —
   `daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || ''` and
   `daysToDeadline: row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] || ''`.
   A grievance with 0 days open (filed today) or 0 days to deadline
   (deadline is today) would lose its data.

2. **Grievance export functions** in `12_Features.gs` (lines 2802–2803,
   3313–3314) — `row[cols.DAYS_OPEN - 1] || ''` and
   `daysToDeadline || ''` in export data arrays. Same zero-erasure bug
   in exported data.

3. **Test helper functions** in `09_Dashboards.test.js` (lines 165–166)
   and `04e_PublicDashboard.test.js` (lines 95–96) — test row builders
   used `overrides.daysOpen || ''` which would silently convert a `0`
   test override to `''`.

**Note:** Other files (e.g. `04c_InteractiveDashboard.gs`) already used
the correct `|| 0` pattern for these same fields.

**Files affected:**
- `01_Core.gs` — `mapGrievanceRow()` now uses `numericField_()`
- `12_Features.gs` — export functions now use `numericField_()`
- `09_Dashboards.test.js` — test helper uses `!= null` ternary
- `04e_PublicDashboard.test.js` — test helper uses `!= null` ternary
- `dist/ConsolidatedDashboard.gs` — consolidated build reflecting all changes

**Resolution:**
All numeric fields in grievance data access now use `numericField_()`
(production code) or null-safe ternaries (test code) to preserve `0`.

---

## Items Confirmed Safe (not bugs)

These were reviewed and determined to be acceptable:

- **Hidden calc sheet column letters** (`_Grievance_Formulas`,
  `_Checklist_Calc`) — These sheets are code-generated with a fixed
  internal layout. Users never interact with them. Hardcoded references
  to their own internal columns are acceptable.
- **Hidden calc sheet hardcoded indices** (`_Member_Lookup`,
  `_Steward_Performance_Calc`) — Same as above. Code-generated hidden
  sheets with fixed layouts. `STEWARD_PERF_COLS` constants exist and
  are used elsewhere; the few hardcoded reads in `09_Dashboards.gs`
  (lines 2070–2079, 2883–2889) are consistent with the sheet layout.
- **Single-column `getRange()` reads with `row[0]`** — When reading a
  1-column-wide range, `row[0]` is the only element. This is correct.
- **External integration sheets** (Volunteer Hours, Meeting Attendance)
  — These don't have defined column constants; hardcoded indices with
  comments are acceptable for now.
- **`SATISFACTION_COLS.AVG_SCHEDULING || 82`** — Defensive fallback
  pattern, not a hardcoded value.
- **Satisfaction survey hardcoded indices** in `04e_PublicDashboard.gs`
  (lines 797–870) — ~50 hardcoded 0-indexed array accesses for survey
  questions (e.g. `satData[i][6]` for Q6). Named constants
  (`SATISFACTION_COLS`) exist and the values match. Not a bug today,
  but a maintainability risk if survey columns are ever reordered.
  Acceptable because survey column order is fixed by Google Forms.
- **`GRIEVANCE_COLUMNS.DRIVE_FOLDER + 1`** in `05_Integrations.gs`
  (lines 352, 371, 380) — Uses 0-indexed legacy `GRIEVANCE_COLUMNS`
  constant with `+ 1` to convert for 1-indexed `getRange()`. Math is
  correct. Could be simplified to `GRIEVANCE_COLS.DRIVE_FOLDER_URL`
  but not a bug.
- **Audit log `getRange(row, 1)`** in `08d_AuditAndFormulas.gs`
  (line 131) — Reads column A (record ID) for both Member Directory
  and Grievance Log. Column A is the ID column in both sheets by
  design, so the hardcoded `1` is correct.

---

## Architecture After Fixes

```
Header Maps (MEMBER_HEADER_MAP_, GRIEVANCE_HEADER_MAP_, etc.)
    |
    v
buildColsFromMap_() -----> *_COLS (1-indexed constants)
    |                           |
    |                           +--> buildLegacyCols_() --> *_COLUMNS (0-indexed)
    |                           +--> buildDropdownMap_() --> DROPDOWN_MAP
    |                           +--> buildMultiSelectCols_() --> MULTI_SELECT_COLS
    |
    v
syncColumnMaps()  <-- called in onOpen()
    |  Reads actual sheet headers at runtime
    |  Updates *_COLS if columns moved
    |  Rebuilds all derived objects
    |  Persists to CacheService (6hr TTL)
    |
    v
loadCachedColumnMaps_()  <-- called in onEdit()
    Restores positions from cache
    Rebuilds derived objects if changed
    Falls back to defaults if cache empty
```

### Single Source of Truth Locations

| What | Where | Used By |
|---|---|---|
| Column order & names | `*_HEADER_MAP_` arrays in `01_Core.gs` | Sheet creation, syncColumnMaps |
| 1-indexed positions | `*_COLS` objects in `01_Core.gs` | All getRange/formula code |
| 0-indexed positions | `*_COLUMNS` objects in `01_Core.gs` | Legacy array access code |
| Dropdown mappings | `DROPDOWN_MAP` in `01_Core.gs` | setupDataValidations, syncDropdownToConfig_ |
| Multi-select mappings | `MULTI_SELECT_COLS` in `01_Core.gs` | setupDataValidations, multi-select dialog |

### Adding a New Column (Checklist)

1. Add entry to the appropriate `*_HEADER_MAP_` array in `01_Core.gs`
2. If it's a dropdown: add entry to `buildDropdownMap_()` or `buildMultiSelectCols_()`
3. Done. Everything else (validation, sync, creation, caching) updates automatically.
