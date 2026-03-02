# SYNC-LOG.md — DDS-Dashboard → Union-Tracker

## Purpose
Tracks all syncs from DDS-Dashboard `Main` to Union-Tracker `Main`.

---

## Definitions

### Data Types
- **System-generated data**: Any data written by code functions (auto-IDs, formulas, computed fields, dashboard KPIs, timestamps). Exception: the import function — its output is treated as manually entered.
- **Manually entered data**: Anything imported by users via the import function, or typed directly into cells by a user. **Must never be deleted or overwritten by any function.**

### Sync Rules
- Source: DDS `Main` → Target: UT `Main`
- **Single branch policy: Both repos use `Main` only.**
- **Excluded from UT:**
  - DDS Apps Script ID (`18hHHX-4E_ykGCqu_EDwKCwqY9ycyRgPtOmguacsxnVZ4YsRh-YETODiu`)
  - Workload Tracker-exclusive files (see registry below)
  - Any credentials, tokens, or secrets (UT is public)

---

## Workload Tracker Exclusion Registry

### ❌ EXCLUDE from UT
| File | Reason |
|------|--------|
| `src/18_WorkloadTracker.gs` | Core WT module — all 32 WT functions |
| `src/WorkloadTracker.html` | WT portal HTML template |

### ⚠️ MODIFY for UT (add typeof guards)
| File | Lines | What to guard |
|------|-------|---------------|
| `src/03_UIComponents.gs` | 85-94 | `📊 Workload Tracker` submenu — wrap in typeof check |
| `src/index.html` | 370, 389 | `workload` nav item and route — guard or remove |
| `src/member_view.html` | 792, 881-1129 | `renderWorkloadTracker` function + SSO calls — guard |

### ✅ KEEP in UT (already safe)
| File | Why safe |
|------|----------|
| `src/05_Integrations.gs` | Already has `typeof getWorkloadTrackerPortalHtml === 'function'` guard |
| `src/08a_SheetSetup.gs` | Already has `typeof initWorkloadTrackerSheets === 'function'` guard |
| `src/01_Core.gs` | Sheet constants defined but unused without WT module — inert |
| `src/styles.html` | WT CSS classes are dead code without module — no errors |
| All other files | Reference "steward workload" (case counts) not the WT module |

### Important Distinction
- **"Workload Tracker"** = member caseload reporting module (18_WorkloadTracker.gs) → EXCLUDE
- **"Steward workload"** = grievance case count per steward → KEEP (core functionality)

---

## Sync History

| Date | Agent | DDS Commit | UT Commit | Files Synced | Exclusions Applied | Notes |
|------|-------|------------|-----------|-------------|-------------------|-------|
| 2026-02-25 | Claude (claude.ai) | — | — | SYNC-LOG.md, AI-REFERENCE.md | N/A | Initial setup files |
| 2026-02-28 | Claude (claude-code) | `d1e51fb` | `b119401` | 102 files (40 .gs, 8 .html, tests, build) | 18_WorkloadTracker.gs, WorkloadTracker.html excluded; typeof guards applied | Full sync: Batches 1-10 code review fixes |
| 2026-03-02 | Claude (claude-code) | `efffbfd` | `d675259` | 22 files (10 src + 11 dist + CHANGELOG.md + package.json) | 18_WorkloadTracker.gs, WorkloadTracker.html excluded; typeof guards preserved in index.html, member_view.html | v4.19.0 sync: QA bug fixes & resilience (Issues 1-12), sign-out fix, security hardening (rate limiting, token cleanup, timing-attack prevention) |

