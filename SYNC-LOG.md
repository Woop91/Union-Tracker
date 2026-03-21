# SYNC-LOG.md — SolidBase

## Purpose
Tracks sync history for the SolidBase repository.

---

## Definitions

### Data Types
- **System-generated data**: Any data written by code functions (auto-IDs, formulas, computed fields, dashboard KPIs, timestamps). Exception: the import function — its output is treated as manually entered.
- **Manually entered data**: Anything imported by users via the import function, or typed directly into cells by a user. **Must never be deleted or overwritten by any function.**

### Sync Rules
- **Single branch policy: `Main` only.**
- **Excluded from public repo:**
  - Any credentials, tokens, or secrets

---

## Workload Tracker Exclusion Registry

> **Resolved (v4.20.0):** The standalone Workload Tracker portal (`18_WorkloadTracker.gs` + `WorkloadTracker.html`) has been deleted. The workload tracker is now fully integrated into the SPA via `25_WorkloadService.gs` and the workload tab in `member_view.html`. No exclusions remain.

### Previously Excluded (now deleted)
| File | Status |
|------|--------|
| `src/18_WorkloadTracker.gs` | **Deleted** — replaced by `25_WorkloadService.gs` |
| `src/WorkloadTracker.html` | **Deleted** — replaced by SPA workload tab in `member_view.html` |

### typeof Guards (kept as defensive coding)
| File | Guard | Notes |
|------|-------|-------|
| `src/index.html` | `typeof` checks for workload nav | Safe — functions exist via `25_WorkloadService.gs` |
| `src/member_view.html` | `typeof` checks for SSO wrappers | Safe — functions exist via `25_WorkloadService.gs` |
| `src/03_UIComponents.gs` | `typeof initWorkloadTrackerSheets` guard in Tools menu | Safe — function exists |

### Important Distinction
- **"Workload Tracker"** = member caseload reporting (now via `25_WorkloadService.gs`)
- **"Steward workload"** = grievance case count per steward → KEEP (core functionality)

---

## Sync History

| Date | Agent | Commit | Files Synced | Notes |
|------|-------|--------|-------------|-------|
| 2026-02-25 | Claude (claude.ai) | — | SYNC-LOG.md, AI-REFERENCE.md | Initial setup files |
| 2026-02-28 | Claude (claude-code) | `b119401` | 102 files (40 .gs, 8 .html, tests, build) | Full sync: Batches 1-10 code review fixes |
| 2026-03-02 | Claude (claude-code) | `d675259` | 22 files (10 src + 11 dist + CHANGELOG.md + package.json) | v4.19.0 sync: QA bug fixes & resilience |
| 2026-03-07 | Claude (claude-code) | `609edc9` | 42 .gs + 7 .html (full parity) | v4.24.4 full sync: Q&A Forum, Timeline, dynamic surveys, auth sweep |
