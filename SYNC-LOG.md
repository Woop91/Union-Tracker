# SYNC-LOG.md — DDS-Dashboard → Union-Tracker

## Purpose
Tracks all syncs from DDS-Dashboard `Main` to Union-Tracker `staging`.

---

## Definitions

### Data Types
- **System-generated data**: Any data written by code functions (auto-IDs, formulas, computed fields, dashboard KPIs, timestamps). Exception: the import function — its output is treated as manually entered.
- **Manually entered data**: Anything imported by users via the import function, or typed directly into cells by a user. **Must never be deleted or overwritten by any function.**

### Sync Rules
- Source: DDS `Main` → Target: UT `staging`
- User manages: UT `staging` → UT `dev` → UT `main`
- **Excluded from UT:**
  - DDS Apps Script ID (`18hHHX-4E_ykGCqu_EDwKCwqY9ycyRgPtOmguacsxnVZ4YsRh-YETODiu`)
  - Workload Tracker-exclusive files (see registry below)
  - Any credentials, tokens, or secrets (UT is public)

---

## Workload Tracker Exclusion Registry

> **Resolved (v4.20.0):** The standalone Workload Tracker portal (`18_WorkloadTracker.gs` + `WorkloadTracker.html`) has been deleted from DDS. The workload tracker is now fully integrated into the SPA via `25_WorkloadService.gs` and the workload tab in `member_view.html`. Both repos are identical — no exclusions remain.

### Previously Excluded (now deleted from both repos)
| File | Status |
|------|--------|
| `src/18_WorkloadTracker.gs` | **Deleted** — replaced by `25_WorkloadService.gs` |
| `src/WorkloadTracker.html` | **Deleted** — replaced by SPA workload tab in `member_view.html` |

### typeof Guards (kept as defensive coding)
| File | Guard | Notes |
|------|-------|-------|
| `src/index.html` | `typeof` checks for workload nav | Safe — functions exist in both repos via `25_WorkloadService.gs` |
| `src/member_view.html` | `typeof` checks for SSO wrappers | Safe — functions exist in both repos via `25_WorkloadService.gs` |
| `src/03_UIComponents.gs` | `typeof initWorkloadTrackerSheets` guard in Tools menu | Safe — function exists in both repos |

### Important Distinction
- **"Workload Tracker"** = member caseload reporting (now via `25_WorkloadService.gs`) → identical in both repos
- **"Steward workload"** = grievance case count per steward → KEEP (core functionality)

---

## Sync History

| Date | Agent | DDS Commit | UT Commit | Files Synced | Exclusions Applied | Notes |
|------|-------|------------|-----------|-------------|-------------------|-------|
| 2026-02-25 | Claude (claude.ai) | — | — | SYNC-LOG.md, AI-REFERENCE.md | N/A | Initial setup files |
| 2026-02-28 | Claude (claude-code) | `d1e51fb` | `b119401` | 102 files (40 .gs, 8 .html, tests, build) | 18_WorkloadTracker.gs, WorkloadTracker.html excluded; typeof guards on 03_UIComponents.gs, index.html, member_view.html; build.js BUILD_ORDER updated; architecture.test.js updated | Full sync: Batches 1-10 code review fixes (auth, XSS, formula injection, perf, dead code) |

