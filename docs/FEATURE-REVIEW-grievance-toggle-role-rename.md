# 🔍 SolidBase — Feature Review: Grievance Toggle + Role Renaming

> **Created:** 2026-03-22  
> **Status:** Proposed — Pending Implementation  
> **Author:** AI Review (Claude)

---

## 📋 Scope Summary

Two independent but architecturally related features:

| Feature | Complexity | Risk |
|---|---|---|
| Grievance visibility toggle | Medium | Low — config-driven flag |
| Steward/Member label renaming | Low | Very Low — text substitution |

Both live entirely in the **Config tab** — zero hardcoding, consistent with the core design principle.

---

## ✅ Independence Confirmed

The two features share the same architectural pattern (Config tab rows + single-pass DOM sweep at boot) but are **entirely separate systems** with zero coupling.

| | Grievance Toggle | Role Rename |
|---|---|---|
| **Config key** | `SHOW_GRIEVANCES` | `ROLE_PRIMARY_LABEL` / `ROLE_SECONDARY_LABEL` |
| **DOM attribute** | `data-feature="grievance"` | `data-label="primary-role"` / `data-label="secondary-role"` |
| **What it does** | Removes/restores elements | Replaces text content |
| **Backend impact** | Yes — guards on GAS functions | No — labels never reach backend |
| **Depends on each other** | ❌ No | ❌ No |

You can:
- Rename roles **without** touching grievance visibility
- Hide grievances **without** renaming anything
- Do both at once — they run as separate passes at init, neither aware of the other

---

## 🔴 Feature 1 — Grievance Visibility Toggle

### What Needs Hiding

| Layer | What Gets Hidden |
|---|---|
| **Navigation** | Grievance nav links in steward & member sidebars |
| **Dashboard KPIs** | Open grievances, overdue, due-this-week counts |
| **WorkloadTracker** | Entire steward workload view (grievance-driven) |
| **Sheets** | Grievance Log sheet tab (hidden, not deleted) |
| **Forms/Modals** | Submit grievance button, grievance detail modals |
| **Data endpoints** | `getGrievances()`, `getWorkload()` server-side functions |
| **Executive Dashboard** | Any Chart.js charts rendering grievance data |

### Recommended Architecture

```
Config Tab
├── Row: SHOW_GRIEVANCES | TRUE / FALSE
```

**`ConfigReader.gs`** — expose it:
```javascript
getShowGrievances() → boolean
```

**Frontend (SPA `index.html`)** — reads flag at boot:
```javascript
// On app init, after config load:
if (!config.showGrievances) {
  document.querySelectorAll('[data-feature="grievance"]').forEach(el => el.remove());
}
```

All grievance UI elements get a single attribute: `data-feature="grievance"` — one selector to rule them all.

**Backend** — guard every grievance endpoint:
```javascript
function getGrievances() {
  if (!ConfigReader.getShowGrievances()) return { error: 'Feature disabled' };
  // ...
}
```

### Toggle Behavior Decisions

| Question | Option A | Option B | Recommendation |
|---|---|---|---|
| What hides the Sheet tab? | Script hides on toggle | Manual | **Script** — use `sheet.hideSheet()` |
| When does toggle apply? | Immediately on config change | On next page load | **Next page load** — Apps Script limit |
| Who can toggle? | Anyone with Config access | Admin-only | **Config tab protection** (existing role model) |
| Partial hide? | All-or-nothing | Per-section | **All-or-nothing** — simpler, safer |

### Implementation Flow

```
Config Tab: SHOW_GRIEVANCES = FALSE
        ↓
ConfigReader.getShowGrievances() → false
        ↓
    ┌──────────────────────────────────────────────────────┐
    │ Frontend: removes [data-feature="grievance"] elements │
    │ Backend: returns disabled errors on grievance calls   │
    │ Sheet: hideSheet() on Grievance Log                   │
    └──────────────────────────────────────────────────────┘
```

---

## 🔵 Feature 2 — Steward / Member Label Renaming

### What Gets Renamed

- Navigation headers
- Role badges / chips
- Dashboard section titles ("Steward View", "Member Portal")
- Form labels ("Submit to your Steward")
- Email templates (if any)
- Sheet tab names (optional — see below)

### Recommended Architecture

```
Config Tab
├── ROLE_PRIMARY_LABEL    | Steward     (default)
├── ROLE_SECONDARY_LABEL  | Member      (default)
```

**`ConfigReader.gs`:**
```javascript
getRoleLabels() → { primary: string, secondary: string }
```

**Frontend — single substitution pass at boot:**
```javascript
const labels = config.roleLabels; // { primary: "Steward", secondary: "Member" }

document.querySelectorAll('[data-label="primary-role"]')
  .forEach(el => el.textContent = labels.primary);

document.querySelectorAll('[data-label="secondary-role"]')
  .forEach(el => el.textContent = labels.secondary);
```

### Sheet Tab Renaming

| Option | Pros | Cons | Recommendation |
|---|---|---|---|
| Rename sheet tabs dynamically | Fully consistent | `sheet.setName()` is destructive — risky | ❌ |
| Leave sheet tabs as-is | Safe, no data risk | Internal naming ≠ UI labels | ✅ Default |
| Separate config row for sheet names | Controlled, explicit | Two separate rename configs | Optional later |

**Decision:** Keep sheet tab names separate from UI labels. Renaming sheets mid-operation risks breaking all codebase references. UI labels only.

---

## 🏗️ Combined Config Tab Design

```
Key                   | Default Value | Notes
----------------------|---------------|---------------------------
SHOW_GRIEVANCES       | TRUE          | FALSE hides all grievance UI + data
ROLE_PRIMARY_LABEL    | Steward       | Displayed wherever steward role appears
ROLE_SECONDARY_LABEL  | Member        | Displayed wherever member role appears
```

Three rows. No deploy required to change behavior — edit Config tab, reload app.

---

## ⚡ Implementation Plan

### Phase 1 — Config + ConfigReader (30 min)
- Add 3 Config rows with defaults
- Extend `ConfigReader` with `getShowGrievances()` and `getRoleLabels()`
- Add fallback defaults so existing deployments don't break

### Phase 2 — Backend guards (1 hr)
- Add `getShowGrievances()` guard to every grievance server function
- Add `sheet.hideSheet()` / `showSheet()` logic tied to config read

### Phase 3 — Frontend attribute tagging (2–3 hrs)
- Tag every grievance UI element: `data-feature="grievance"`
- Tag every role label: `data-label="primary-role"` or `data-label="secondary-role"`
- Wire both init sweeps in app boot sequence

### Phase 4 — Testing
- Toggle OFF → no grievance UI leaks, no JS errors from missing elements
- Label rename → every surface updates, no hardcoded strings remain
- Toggle back ON → full restore verified

---

## ❓ Unresolved Questions

1. Toggle OFF — hide Grievance Log sheet tab, or leave visible to admins?
2. Sheet tab names: rename dynamically later, or UI labels only permanently?
3. "Primary/Secondary" label naming accurate, or different terminology needed?
4. Should toggle state changes be logged in audit log?
5. SolidBase-only change, or sync equivalent config keys to SolidBase too?
