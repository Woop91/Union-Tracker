# Sub-Project D — Admin UX Enhancements

**Date:** 2026-04-11
**Target release:** DDS v4.55.6 (minor enhancement wave)
**Status:** Compact spec, batch-approval mode
**Closes requests:**
- "Increase the size of the org name in the top left corner of the side bar in the web app to be larger, take more of the vertical space"
- "If there are no notifications available, the notification bell should not be visible"
- "Add DataFailsafe and TestRunner to the admin tab"
- "Can the usage analytics track browsers used. iPhone, Android usage. Are there other things I might want to track as admin?"

---

## 1. Motivation

Four small admin-facing improvements landed in one batch because they share a single file surface (sidebar rendering and admin settings) and none of them alone justifies a release cycle. The biggest is the usage analytics expansion (browser/OS/device parsing). The others are UX polish.

## 2. Scope

Four zones:

1. **Zone 1 — Larger sidebar org name.** CSS change in `src/styles.html:70-97`. Bump `.sidebar-header-title` from `1rem` (16px) to `1.375rem` (22px), increase `line-height`, tweak padding so the header block takes more vertical space. No DOM change.

2. **Zone 2 — Hide notification bell when empty.** In `src/index.html:2800-2809`, the bell wrapper is always created and appended, and only the *badge* is conditional. Wrap the entire `appendChild(bellWrap)` in the existing `if (_bellTotal > 0)` check so the bell itself is absent when there are no notifications. Keep the click handler logic intact for the non-zero case.

3. **Zone 3 — Promote DataFailsafe + TestRunner to Admin group.** Per investigation, these two tabs already exist in `_getSidebarTabs('steward')` at `src/index.html:3197-3198` under a "Temp" group (dev-only, likely filtered out in prod via `_filterDevTabs()`). Move them from "Temp" to the "Admin" group that's added for `IS_ADMIN` users (line 3201-3205). Verify the render functions exist — if they do, Zone 3 is a 4-line array shuffle. If they don't, create minimal stubs that explain "feature coming soon" rather than shipping broken tabs. Backend wrappers (`dataRunFailsafe`, `dataRunTests`, etc.) are deferred unless rendering requires them.

4. **Zone 4 — Usage analytics browser/OS/device parsing.** Currently `dataLogUsageEvents` captures `navigator.userAgent` as a raw string and the aggregation in `dataGetUsageAnalytics` does a crude regex match for mobile/tablet/desktop. Add a small `parseUserAgent(ua)` helper (client-side, before sending) that extracts:
   - `browser`: one of `chrome`, `safari`, `firefox`, `edge`, `other`
   - `os`: one of `ios`, `android`, `macos`, `windows`, `linux`, `other`
   - `deviceClass`: one of `mobile`, `tablet`, `desktop` (existing logic, kept)
   
   Send these as additional fields in the event payload. Extend `dataGetUsageAnalytics` aggregation to group by browser + OS as well as device class. Extend the Analytics admin tab render to show the new groupings.

## 3. Architecture & fix shapes

### Zone 1 (CSS)

`src/styles.html:70-97` currently:
```css
.sidebar-header { padding: 12px 20px 20px; }
.sidebar-header-title { font-size: 1rem; font-weight: 700; }
.sidebar-header-sub { font-size: 0.6875rem; }
```

Change to:
```css
.sidebar-header { padding: 16px 20px 24px; }
.sidebar-header-title { font-size: 1.375rem; font-weight: 800; line-height: 1.15; letter-spacing: -0.01em; }
.sidebar-header-sub { font-size: 0.75rem; margin-top: 4px; }
```

Effect: org name jumps from 16px to 22px, sits on two lines more comfortably if long, header block grows by ~12-16px vertically.

### Zone 2 (bell hide)

`src/index.html:2800-2809` currently (simplified):
```javascript
var bellWrap = el('div', { className: 'notif-bell-wrap', ... });
bellWrap.appendChild(el('span', { className: 'notif-bell-icon' }, '\uD83D\uDD14'));
var _bellTotal = (AppState.notificationCount || 0) + (role === 'steward' ? (AppState.qaUnansweredCount || 0) : 0);
if (_bellTotal > 0) {
  bellWrap.appendChild(el('span', { className: 'notif-bell-badge', ... }, String(_bellTotal)));
}
header.appendChild(bellWrap);
```

Change to:
```javascript
var _bellTotal = (AppState.notificationCount || 0) + (role === 'steward' ? (AppState.qaUnansweredCount || 0) : 0);
if (_bellTotal > 0) {
  var bellWrap = el('div', { className: 'notif-bell-wrap', ... });
  bellWrap.appendChild(el('span', { className: 'notif-bell-icon' }, '\uD83D\uDD14'));
  bellWrap.appendChild(el('span', { className: 'notif-bell-badge', ... }, String(_bellTotal)));
  header.appendChild(bellWrap);
}
```

Reorders so bell creation is inside the `if`. If `_bellTotal === 0`, no bell is added to the DOM at all.

### Zone 3 (tab group move)

`src/index.html:3169-3209` contains the hardcoded steward tab array. The two tabs are at ~3197-3198 under a `group: 'Temp'` section. The admin-only group is added at ~3201-3205 conditionally when `IS_ADMIN`.

Change:
- Remove `{ id: 'failsafe', ... }` and `{ id: 'testrunner', ... }` from the Temp group.
- Add them to the admin-only block, so they only appear for admin users.

```javascript
if (IS_ADMIN) {
  stewardTabs.push({ divider: true });
  stewardTabs.push({ group: 'Admin' });
  stewardTabs.push({ id: 'analytics', icon: '📊', label: 'Usage Analytics' });
  stewardTabs.push({ id: 'failsafe', icon: '🛡️', label: 'Data Failsafe' });
  stewardTabs.push({ id: 'testrunner', icon: '🧪', label: 'Test Runner' });
}
```

Verify the tab router at `_handleTabNav` has cases for `'failsafe'` and `'testrunner'`. If the render functions exist, no further change. If they don't, create minimal "stub" pages that render a 1-line explainer and a single button that calls the backend. **Decision gate during implementation.**

### Zone 4 (userAgent parsing)

**New helper** in `src/index.html` (inside the shared IIFE, near `renderSafeText` / `displayUser` added in sub-project A):

```javascript
function parseUserAgent(ua) {
  ua = String(ua || '');
  var lua = ua.toLowerCase();
  var browser = 'other';
  if (lua.indexOf('edg/') !== -1) browser = 'edge';
  else if (lua.indexOf('chrome/') !== -1 && lua.indexOf('edg/') === -1) browser = 'chrome';
  else if (lua.indexOf('firefox/') !== -1) browser = 'firefox';
  else if (lua.indexOf('safari/') !== -1 && lua.indexOf('chrome/') === -1) browser = 'safari';
  var os = 'other';
  if (/iphone|ipad|ipod/i.test(ua)) os = 'ios';
  else if (/android/i.test(ua)) os = 'android';
  else if (/macintosh|mac os/i.test(ua)) os = 'macos';
  else if (/windows/i.test(ua)) os = 'windows';
  else if (/linux/i.test(ua)) os = 'linux';
  var deviceClass = 'desktop';
  if (/mobile|android|iphone/i.test(ua)) deviceClass = 'mobile';
  else if (/ipad|tablet/i.test(ua)) deviceClass = 'tablet';
  return { browser: browser, os: os, deviceClass: deviceClass };
}
```

Call this at the start of `logUsageEvent` (or whatever local wrapper calls `dataLogUsageEvents`) and add `browser`, `os`, `deviceClass` to the event payload.

**Backend change** in `src/21d_WebDashDataWrappers.gs`'s `dataLogUsageEvents`: add three new columns (`Browser`, `OS`, `Device Class`) to the `_Usage_Log` sheet header on init, and write the parsed values from the payload.

**Aggregation change** in `dataGetUsageAnalytics`: group by `browser` and `os` in addition to the existing `deviceTypes` breakdown. Return two new arrays: `byBrowser` and `byOs`.

**Admin render change**: in the Usage Analytics admin page render (find via grep — likely in `steward_view.html` or a dedicated analytics file), add two new sections showing the browser and OS breakdowns. Simple text rows are enough; no new charts.

## 4. Other metrics worth tracking as admin (answering the user's open question)

Proposed but NOT implemented in this wave (defer to a future D+ wave if you want them):
- **Session length distribution** — histogram of session durations (< 1min, 1-5min, 5-15min, 15-30min, 30min+). Gives you "how many users actually use the app vs bounce."
- **Error rate per role** — already tracked as events, but not aggregated per role. Easy addition.
- **Top 10 grievance categories filed** — from the existing grievance log, not from usage analytics. Would need a separate endpoint.
- **Dues-paying vs non-paying usage split** — which features do non-dues members actually use?
- **Cold-start vs warm-start performance** — you already track `perf_load` and `perf_batch`; differentiate by whether it's the user's first session of the day.

All of these are 1-3 hour additions. This spec includes only the browser/OS/device parse because it matches the user's literal request. The rest should be brainstormed separately if wanted.

## 5. Test plan

~12 new tests:
- `test/parse-user-agent.test.js` (~8 tests): iPhone, iPad, Android, Windows Chrome, Windows Edge, macOS Safari, macOS Firefox, Linux Firefox — each asserts the correct `{browser, os, deviceClass}` output.
- `test/spa-integrity.test.js` additions (~4): G-ADMIN-TABS — assert that when `IS_ADMIN` is true, the `failsafe`, `testrunner`, and `analytics` tabs are in the admin group; assert notif-bell append is inside the `_bellTotal > 0` guard; assert `.sidebar-header-title` font-size is larger than `1rem`.

## 6. Deploy

Per-zone commits, version bump v4.55.5 → v4.55.6, push + clasp DDS prod. SolidBase sync deferred per Option Y.

## 7. Success criteria

- Sidebar org name visibly larger (≥22px) in both light and dark themes.
- Notification bell completely absent from the DOM when `_bellTotal === 0`.
- Admin users see `Data Failsafe` and `Test Runner` tabs in the Admin group alongside `Usage Analytics`.
- Usage analytics captures `browser`, `os`, `deviceClass` per event and the Analytics page shows the breakdown.
- DDS v4.55.6 live on prod.
