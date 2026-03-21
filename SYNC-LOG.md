# SolidBase Sync Log

## Sync Policy
- **Source of truth:** DDS-Dashboard Main
- **Target:** SolidBase Main
- **Exclusions from SolidBase:**
  - DDS Apps Script ID (redacted)
  - `25_WorkloadService.gs` + `poms_reference.html` (org-specific features)
  - `25_WorkloadService.test.js`, `scripts/sync-org-chart.js`
  - Org-specific references (DDS, MassAbility, SEIU 509) replaced with generic equivalents
  - `UI_REVIEW.md`, `docs/archived-reviews/` audit docs (DDS-only)
- **SolidBase-specific branding:** "SolidBase — Built for the collective."

## Sync History

### 2026-03-21 — Full sync from DDS v4.33.0
- Source: DDS-Dashboard `e172d7c` (Main)
- Scope: All src/, test/, docs/, scripts/, root configs
- Changes: Full file copy, org reference scrub, POMS + Workload Tracker feature removal
- Files synced: 50+ src, 55+ test, 25+ docs
- Excluded: 25_WorkloadService.gs, poms_reference.html, sync-org-chart.js
