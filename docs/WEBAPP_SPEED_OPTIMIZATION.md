# Webapp Speed Optimization Report

**Date:** 2026-03-14
**Current tab load time:** ~3-4 seconds
**Target:** 1-2 seconds (ideal: 1s)
**Status:** Ready for implementation

---

## Root Cause Summary

The webapp suffers from three compounding issues:

1. **Sequential server round trips** — steward dashboard fires `dataGetStewardCases()` then waits for it to complete before firing `dataGetStewardKPIs()`. Each GAS round trip costs 1-1.5s.
2. **No client-side caching on steward_view** — `member_view.html` uses `DataCache.cachedCall()` (10 usages), but `steward_view.html` has zero. Every tab switch re-fetches all data.
3. **Full DOM destruction on every tab switch** — `renderPageLayout()` nukes the entire page (`container.innerHTML = ''`) including sidebar/nav and rebuilds from scratch.

Combined with expensive CSS (infinite animations, backdrop-filter blur on every card) and no list pagination, these create the 3-4 second delay.

---

## Optimizations — Ordered by Impact

### OPT-1: Create single batch endpoint for steward dashboard init

**Savings: 1.0-1.5s | Effort: Medium | Priority: CRITICAL**

**Problem:**
- `steward_view.html` fires these calls sequentially on dashboard load:
  1. `dataGetStewardCases(SESSION_TOKEN)` — waits for response
  2. `dataGetStewardKPIs(SESSION_TOKEN)` — fires only after cases return
  3. `dataGetBadgeCounts(SESSION_TOKEN)` — separate call
- KPIs are already computed FROM cases on the server side (`_computeKPIsFromCases()`)
- Each GAS `google.script.run` round trip = ~1-1.5 seconds

**Fix:**
1. In `src/21_WebDashDataService.gs`, create a new function:
```javascript
function dataGetStewardDashboardInit(token) {
  var e = _resolveCallerEmail(); if (!e) return { cases: [], kpis: {}, badges: {} };
  var s = _requireStewardAuth(); if (!s) return { cases: [], kpis: {}, badges: {} };
  var cases = getStewardCases(e);
  var kpis = _computeKPIsFromCases(cases);
  var badges = _getBadgeCounts(e);
  return { cases: cases, kpis: kpis, badges: badges };
}
```
2. In `src/steward_view.html`, replace the sequential calls with a single:
```javascript
serverCall('dataGetStewardDashboardInit')
  .withSuccessHandler(function(result) {
    renderCaseList(result.cases);
    renderKPIs(result.kpis);
    updateBadges(result.badges);
  })
  .dataGetStewardDashboardInit(SESSION_TOKEN);
```

**Files to modify:**
- `src/21_WebDashDataService.gs` — add `dataGetStewardDashboardInit()`
- `src/steward_view.html` — refactor steward dashboard init to use single call

---

### OPT-2: Add client-side DataCache to steward_view

**Savings: 1.0-1.5s on revisits | Effort: Low | Priority: CRITICAL**

**Problem:**
- `src/member_view.html` uses `DataCache.cachedCall()` 10 times — steward_view uses it 0 times
- Every time a user navigates away from a steward tab and comes back, ALL data is re-fetched from the server
- The `DataCache` infrastructure already exists in `index.html` (~lines 418-463) with configurable TTLs

**Fix:**
In `src/steward_view.html`, wrap every `serverCall()` with `DataCache.cachedCall()`:
```javascript
// BEFORE (current):
serverCall('dataGetStewardCases')
  .withSuccessHandler(onSuccess)
  .dataGetStewardCases(SESSION_TOKEN);

// AFTER:
DataCache.cachedCall(
  'stewardCases_' + CURRENT_USER.email,
  'dataGetStewardCases',
  [SESSION_TOKEN],
  onSuccess,
  onFailure,
  DataCache.STABLE_TTL  // 5 minutes
);
```

**Key cache TTL guidance:**
- Cases/KPIs: `DataCache.FRESH_TTL` (2 min) — changes frequently
- Members list: `DataCache.STABLE_TTL` (5 min) — changes rarely
- Badge counts: `DataCache.FRESH_TTL` (2 min)
- Config/resources: `DataCache.STABLE_TTL` (5 min)

**After OPT-1 is implemented**, cache the combined batch result from `dataGetStewardDashboardInit` instead of individual calls.

**Files to modify:**
- `src/steward_view.html` — wrap all server calls with DataCache

---

### OPT-3: Seed steward data from preloaded batch

**Savings: 0.5-1.0s on initial load | Effort: Low | Priority: HIGH**

**Problem:**
- `index.html:96-112` fires `dataGetBatchData()` immediately on page load which returns cases, members, tasks, KPIs
- `member_view.html:19-26` properly seeds its local cache from this batch data
- `steward_view.html` completely ignores the preloaded batch and re-fetches everything

**Fix:**
At the top of the steward dashboard init function in `steward_view.html`, check for preloaded data:
```javascript
function initStewardDashboard() {
  _onPreloadReady(function(batch) {
    if (batch && batch.cases) {
      // Seed from preload — no server call needed
      renderCaseList(batch.cases);
      renderKPIs(batch.kpis || _computeKPIsFromCases(batch.cases));
      updateBadges(batch.badges);
      return;
    }
    // Fallback: fetch from server
    serverCall('dataGetStewardDashboardInit')...
  });
}
```

**Files to modify:**
- `src/steward_view.html` — add `_onPreloadReady()` check at steward init

---

### OPT-4: Preserve layout shell on tab switch

**Savings: 0.3-0.5s | Effort: Medium | Priority: HIGH**

**Problem:**
- `renderPageLayout()` at `index.html:842-862` does `container.innerHTML = ''` on EVERY tab switch
- This destroys the sidebar nav, bottom nav, header, theme picker — then recreates all of them
- `renderSidebarItems()` rebuilds the entire sidebar including user name, all nav items, click handlers, theme picker — even though none of this changes between tabs

**Fix:**
1. On first render, create the layout shell (sidebar + content area) and store a reference to the content container
2. On tab switch, only clear and re-render the content area; update the active tab highlight in the sidebar via class toggle
```javascript
var _layoutShell = null;
var _contentContainer = null;

function renderPageLayout(container, role, activeTab, contentFn) {
  if (!_layoutShell) {
    // First render: build full layout
    _layoutShell = el('div', { className: 'page-layout' });
    var sidebar = el('nav', { className: 'sidebar-nav' });
    renderSidebarItems(sidebar, role, activeTab);
    _layoutShell.appendChild(sidebar);
    _contentContainer = el('div', { className: 'page-layout-content' });
    _layoutShell.appendChild(_contentContainer);
    container.appendChild(_layoutShell);
    renderBottomNav(container, role, activeTab);
  }
  // Update active states
  updateActiveTab(activeTab);
  // Only re-render content
  _contentContainer.innerHTML = '';
  contentFn(_contentContainer);
}
```

**Files to modify:**
- `src/index.html` — refactor `renderPageLayout()`, add `updateActiveTab()` helper

---

### OPT-5: Paginate large lists

**Savings: 0.2-0.5s (scales with data) | Effort: Medium | Priority: HIGH**

**Problem:**
- `renderCaseList()` in `steward_view.html` (~line 377-450) creates a DOM card for every single case in a `forEach` loop with no limit
- Each case card creates 6+ child elements + an event listener
- 50 cases = 300+ DOM nodes; 200 cases = 1200+ DOM nodes created synchronously
- Same issue for members list, notifications, tasks, resources

**Fix:**
1. Add a `PAGE_SIZE` constant (e.g., 25)
2. Render only the first page of items
3. Add a "Show more" button that appends the next page
4. Use `DocumentFragment` for batch DOM insertion:
```javascript
var PAGE_SIZE = 25;
function renderCaseList(cases, page) {
  page = page || 0;
  var start = page * PAGE_SIZE;
  var slice = cases.slice(start, start + PAGE_SIZE);
  var frag = document.createDocumentFragment();
  slice.forEach(function(c) {
    var card = el('div', { className: 'card card-clickable' });
    // ... build card ...
    frag.appendChild(card);
  });
  listContainer.appendChild(frag); // single DOM write
  if (start + PAGE_SIZE < cases.length) {
    // Add "Show more" button
  }
}
```

**Files to modify:**
- `src/steward_view.html` — `renderCaseList()`, member list renderer, notification renderer

---

### OPT-6: Batch/parallelize Members tab server calls

**Savings: 0.5-1.0s | Effort: Low | Priority: HIGH**

**Problem:**
- Members tab fires two independent calls sequentially:
  1. `dataGetAllMembers(SESSION_TOKEN)` (~line 757)
  2. `dataGetStewardMemberStats(SESSION_TOKEN)` (~line 669)
- These are independent — neither depends on the other's result

**Fix — Option A (quick):** Fire both in parallel:
```javascript
var membersLoaded = false, statsLoaded = false;
var membersData, statsData;
function tryRender() {
  if (membersLoaded && statsLoaded) renderMembersTab(membersData, statsData);
}
serverCall('dataGetAllMembers').withSuccessHandler(function(d) {
  membersData = d; membersLoaded = true; tryRender();
}).dataGetAllMembers(SESSION_TOKEN);
serverCall('dataGetStewardMemberStats').withSuccessHandler(function(d) {
  statsData = d; statsLoaded = true; tryRender();
}).dataGetStewardMemberStats(SESSION_TOKEN);
```

**Fix — Option B (better):** Create `dataGetMembersTabInit()` that returns both in one call.

**Files to modify:**
- `src/steward_view.html` — Members tab init
- `src/21_WebDashDataService.gs` — (Option B) add combined endpoint

---

### OPT-7: Remove infinite glow-bar CSS animations

**Savings: ~100ms + eliminates scroll jank | Effort: Trivial | Priority: MEDIUM**

**Problem:**
- `styles.html:227-239` — `.glow-bar` runs `glowSweep` animation with `animation: glowSweep 2.5s linear infinite`
- Every grievance card has a glow-bar, so 50 visible cases = 50 concurrent infinite GPU animations
- Causes constant GPU compositing work and scroll jank

**Fix:**
Change from infinite to single iteration, or remove animation entirely:
```css
/* Option A: Run once */
.glow-bar {
  animation: glowSweep 2.5s linear 1;  /* was: infinite */
}

/* Option B: Static gradient (recommended — saves GPU entirely) */
.glow-bar {
  /* Remove animation property entirely */
  background-size: 100% 100%;  /* was: 300px 100% */
}
```

**Files to modify:**
- `src/styles.html` — `.glow-bar` animation rule (~line 232)

---

### OPT-8: Remove backdrop-filter blur from cards

**Savings: 50-150ms render | Effort: Trivial | Priority: MEDIUM**

**Problem:**
- `styles.html:157-158, 299-300, 548` — `backdrop-filter: blur(20px)` on `.card-glass`, modals, and other elements
- Forces GPU compositing layer per element. With many cards visible, this compounds
- Especially slow on low-end devices

**Fix:**
Replace blur with a semi-transparent solid background:
```css
/* BEFORE */
.card-glass {
  backdrop-filter: blur(20px);
  -webkit-backdrop-filter: blur(20px);
  background: var(--card-glass-bg);
}

/* AFTER */
.card-glass {
  /* Remove backdrop-filter lines */
  background: var(--card-bg);  /* Use solid card background */
}
```

Keep `backdrop-filter` ONLY on modals/overlays where the blur effect is most visible and there's only one element.

**Files to modify:**
- `src/styles.html` — remove `backdrop-filter` from `.card-glass` (~line 157), keep on modals only

---

### OPT-9: Use DocumentFragment for all list rendering

**Savings: 50-100ms per list | Effort: Low | Priority: MEDIUM**

**Problem:**
- Case list, member list, notification list, task list all call `container.appendChild(card)` inside a `forEach` loop
- Each `appendChild` on a live DOM node can trigger layout recalculation
- `steward_view.html` has 1,328 `appendChild()` calls across all rendering paths

**Fix:**
Wrap all list rendering in `DocumentFragment`:
```javascript
var frag = document.createDocumentFragment();
items.forEach(function(item) {
  frag.appendChild(buildCard(item));
});
container.appendChild(frag); // single reflow
```

**Files to modify:**
- `src/steward_view.html` — all `forEach` loops that `appendChild` to a live container

---

### OPT-10: Debounce filter/sort re-renders

**Savings: Perceived smoothness | Effort: Trivial | Priority: LOW**

**Problem:**
- Filter dropdowns and sort controls trigger immediate full re-renders of lists
- Rapid changes (e.g., typing in search) cause redundant render cycles

**Fix:**
```javascript
var _filterTimer = null;
function debouncedRerender() {
  clearTimeout(_filterTimer);
  _filterTimer = setTimeout(rerender, 150);
}
sortSelect.addEventListener('change', debouncedRerender);
filterSelect.addEventListener('change', debouncedRerender);
```

**Files to modify:**
- `src/steward_view.html` — filter/sort event handlers

---

## Expected Impact Summary

| # | Optimization | Time Saved | Cumulative |
|---|---|---|---|
| OPT-1 | Single batch endpoint | 1.0-1.5s | ~2.5s |
| OPT-2 | Client-side DataCache | 1.0-1.5s (revisits) | ~1.5s (revisits) |
| OPT-3 | Seed from preloaded batch | 0.5-1.0s | ~1.5s (first load) |
| OPT-4 | Preserve layout shell | 0.3-0.5s | ~1.2s |
| OPT-5 | Paginate large lists | 0.2-0.5s | ~1.0s |
| OPT-6 | Parallel Members tab calls | 0.5-1.0s | ~0.7s |
| OPT-7 | Kill glow-bar animations | ~100ms + jank | ~0.6s |
| OPT-8 | Remove backdrop-filter blur | 50-150ms | ~0.5s |
| OPT-9 | DocumentFragment batching | 50-100ms | ~0.4s |
| OPT-10 | Debounce filters | Perceived | — |

**Implementing OPT-1 through OPT-3** (the three critical items) should reduce tab loads from 3-4s to **~1.5-2s**.

**Adding OPT-4 through OPT-6** should consistently hit the **~1s target**.

**OPT-7 through OPT-10** are polish for perceived smoothness and scroll performance.

## Implementation Order

Recommended sequence to minimize risk and maximize early wins:

1. **OPT-7 + OPT-8** (trivial CSS changes, no logic risk)
2. **OPT-2** (add DataCache to steward_view — infrastructure already exists)
3. **OPT-3** (seed from preloaded batch — pattern already in member_view)
4. **OPT-1** (new batch endpoint — most impactful but needs server + client changes)
5. **OPT-9** (DocumentFragment — mechanical refactor)
6. **OPT-6** (parallelize Members tab)
7. **OPT-4** (layout shell preservation — most structural change)
8. **OPT-5** (pagination — UX change, needs design consideration)
9. **OPT-10** (debounce — trivial)

## Files Requiring Changes

| File | Optimizations |
|------|---------------|
| `src/21_WebDashDataService.gs` | OPT-1, OPT-6 |
| `src/steward_view.html` | OPT-1, OPT-2, OPT-3, OPT-5, OPT-6, OPT-9, OPT-10 |
| `src/index.html` | OPT-4 |
| `src/styles.html` | OPT-7, OPT-8 |

## Testing Notes

- After OPT-1: verify `dataGetStewardDashboardInit` returns correct shape, auth check present
- After OPT-2: verify cache invalidation after mutations (case create/edit/delete must clear cache)
- After OPT-4: verify all tab nav items still highlight correctly, theme picker still works
- After OPT-5: verify all items are still accessible via "Show more", filter/sort work across pages
- Run `npm run test` after each optimization to catch regressions
- Run GAS TestRunner after deploying to verify in-spreadsheet behavior
